"""
emailScraper.py
---------------
Streamlined backend module for scraping Target order and cancellation emails,
with verbose logging so the GUI can display every backend action.

Behavior:
- Connect via IMAP and parse retail order emails.
- Maintain a per-account JSON cache to avoid reprocessing.
- Process cached emails and update the orders CSV (_orders.csv) so that
  order statuses advance in the logical shipping order: ordered -> shipped -> delivered.
  Delivered is terminal (no further advances).
- Cancellations are written to _cancellations.csv and orders marked as CANCELLED.
- Verbose prints/logs added at key steps so the GUI receives real-time feedback.
"""

import imaplib
import email
import re
import os
import csv
import json
from email.header import decode_header
from bs4 import BeautifulSoup
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
from collections import defaultdict
from config import EMAIL_ACCOUNTS, DAYS_BACK, CSV_PATH, CACHE_JSON

# -----------------------
# Retailer rules (data-driven)
# -----------------------
RETAILER_RULES = {
    "target": {
        "ids": ["target", "target.com", "orders@target", "order@target"],
        "order_patterns": [
            r"order\s*#?\s*(\d{8,15})",
            r"order\s*number[:\s]*(\d{8,15})",
            r"#(\d{8,15})"
        ],
        "cancel_indicators": ["cancel", "canceled", "cancelled", "sorry, we had to cancel"],
        "shipped_indicators": ["shipped", "your item is on the way", "your package"],
        "delivered_indicators": ["delivered", "out for delivery", "arrived", "left at the", "was delivered"],
    },
    "walmart": {
        "ids": ["walmart", "walmart.com"],
        "order_patterns": [
            r"\b\d{6,10}-\d{6,12}\b",
            r"order\s*number[:\s#]*([\d-]{10,25})",
            r"order\s*#?\s*([\d-]{10,25})"
        ],
        "cancel_indicators": ["canceled", "cancelled", "canceled:"],
        "shipped_indicators": ["shipped", "your package shipped"],
        "delivered_indicators": ["arrived", "your package arrived", "delivered"],
        "tracking_patterns": [r"\b1Z[0-9A-Z]{16}\b", r"\b(\d{12,22})\b"]
    },
    "amazon": {
        "ids": ["amazon", "amazon.com"],
        "order_patterns": [
            r"\b(\d{3}-\d{7}-\d{7})\b",
            r"order\s*#?\s*(\d{3}-\d{7}-\d{7})"
        ],
        "cancel_indicators": ["canceled", "cancelled", "order canceled", "order cancelled"],
        "shipped_indicators": ["shipped", "out for delivery", "your package is on the way"],
        "delivered_indicators": ["delivered", "has been delivered", "arrived"],
    },
    "bestbuy": {
        "ids": ["bestbuy", "best buy"],
        "order_patterns": [
            r"order\s*#?\s*(\d{6,15})",
            r"order\s*number[:\s]*(\d{6,15})"
        ],
        "cancel_indicators": ["canceled", "cancelled"],
        "shipped_indicators": ["shipped"],
        "delivered_indicators": ["delivered", "arrived"],
    }
}

GLOBAL_TRACKING_PATTERNS = [
    r"\b(1Z[0-9A-Z]{16})\b",            # UPS
    r"\b(\d{12,22})\b",                 # FedEx / long numeric
    r"\b(\d{20,22}|\d{13})\b"           # USPS / variations
]


# -----------------------
# Helpers / I/O / Logging
# -----------------------
def log(msg):
    print(f"[LOG] {msg}")


def load_email_cache(filename):
    if not filename:
        log("No cache filename provided to load_email_cache()")
        return []
    try:
        if os.path.exists(filename):
            with open(filename, "r", encoding="utf-8") as f:
                data = json.load(f)
                log(f"Loaded cache from {filename} ({len(data)} entries)")
                return data
        else:
            log(f"Cache file {filename} does not exist (starting fresh).")
    except Exception as e:
        log(f"Could not load cache {filename}: {e}")
    return []


def save_email_cache(cache_data, filename):
    if not filename:
        log("No cache filename provided to save_email_cache()")
        return
    try:
        os.makedirs(os.path.dirname(filename), exist_ok=True)
        with open(filename, "w", encoding="utf-8") as f:
            json.dump(cache_data, f, indent=2)
        log(f"Saved {len(cache_data)} cache entries to {filename}")
    except Exception as e:
        log(f"Could not save cache {filename}: {e}")


def write_to_csv(row, csv_file):
    if not csv_file:
        raise ValueError("csv_file path required")
    os.makedirs(os.path.dirname(csv_file), exist_ok=True)
    file_needs_header = not os.path.isfile(csv_file) or os.stat(csv_file).st_size == 0
    with open(csv_file, "a", newline="", encoding="utf-8") as f:
        if csv_file.endswith("_orders.csv") or csv_file.endswith("report_orders.csv"):
            fieldnames = ["order_number", "tracking_numbers", "ship_to", "sent_to", "sent_date", "status", "retailer"]
        else:
            fieldnames = ["order_number", "sent_to", "sent_date", "reason", "retailer"]
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        if file_needs_header:
            writer.writeheader()
            log(f"Wrote header to {csv_file}")
        writer.writerow(row)
        log(f"Wrote row to {csv_file}: {row.get('order_number')} ({row.get('retailer')})")


# -----------------------
# Parsing helpers
# -----------------------
def clean_subject(subject):
    decoded_parts = decode_header(subject)
    subject_str = ""
    for part, encoding in decoded_parts:
        try:
            if isinstance(part, bytes):
                subject_str += part.decode(encoding or "utf-8", errors="ignore")
            else:
                subject_str += part
        except Exception:
            subject_str += str(part)
    return subject_str


def identify_retailer(sender, subject):
    """
    Identify retailer from sender or subject text (case-insensitive).
    """
    s = (sender or "").lower()
    subj = (subject or "").lower()
    for rname, rules in RETAILER_RULES.items():
        for ident in rules.get("ids", []):
            if ident in s or ident in subj:
                log(f"Identified retailer '{rname}' from sender/subject (ident='{ident}')")
                return rname
    return None


def find_order_number_with_patterns(text, patterns):
    for pat in patterns:
        m = re.search(pat, text or "", re.IGNORECASE)
        if m:
            # if capturing groups present return first group else full match
            return m.group(1) if m.groups() else m.group(0)
    return None


def extract_order_details(text, subject):
    """
    Extract order info using retailer-specific rules. Includes a robust fallback
    for Target: if no explicit pattern matched, try a general 8-15 digit search.
    """
    # Retailer detection: check sender-like strings in the text and subject
    retailer = identify_retailer(text, subject) or identify_retailer(subject, text)

    # Default to 'target' if nothing matched (keeps backward compatibility)
    if retailer is None:
        retailer = "target"

    rules = RETAILER_RULES.get(retailer, RETAILER_RULES["target"])

    # attempt order number detection with retailer patterns (body then subject)
    order_number = find_order_number_with_patterns(text, rules.get("order_patterns", []))
    if not order_number:
        order_number = find_order_number_with_patterns(subject, rules.get("order_patterns", []))

    # --- TARGET fallback: look for any 8-15 digit sequence if usual patterns missed ---
    if retailer == "target" and not order_number:
        # If the rules didn't find it, be more flexible: any 8-15 digit sequence is likely a Target order #
        m = re.search(r"\b(\d{8,15})\b", text or "") or re.search(r"\b(\d{8,15})\b", subject or "")
        if m:
            order_number = m.group(1)
            log(f"Target fallback matched order number {order_number}")

    # Tracking detection
    tracking_numbers = []
    tracking_patterns = rules.get("tracking_patterns") or GLOBAL_TRACKING_PATTERNS
    for tpat in tracking_patterns:
        found = re.findall(tpat, text or "")
        if found:
            flat = []
            for f in found:
                if isinstance(f, tuple):
                    flat.append(next((x for x in f if x), ""))
                else:
                    flat.append(f)
            tracking_numbers.extend(flat)
    tracking_numbers = list({t for t in tracking_numbers if t and t != order_number})

    # Ship_to extraction (simple)
    address_match = re.search(r'Delivers to:\s*(.+,\s*\d{5})', text or "", re.IGNORECASE)
    ship_to = address_match.group(1).strip() if address_match else None

    # Cancellation detection
    is_cancel = False
    for kw in rules.get("cancel_indicators", []):
        if kw in (subject or "").lower() or kw in (text or "").lower():
            is_cancel = True
            break

    cancellation_reason = None
    if is_cancel:
        reason_patterns = [
            r'what went wrong\?\s*(.+?)(?:\*\*|$)',
            r'reason[:\s]*([^.\n]+)',
            r'because[:\s]*([^.\n]+)',
            r'unfortunately[:\s]*([^.\n]+)',
        ]
        for pat in reason_patterns:
            rm = re.search(pat, text or "", re.IGNORECASE | re.DOTALL)
            if rm:
                cancellation_reason = rm.group(1).strip()[:200]
                break
        if not cancellation_reason:
            cancellation_reason = "Reason not specified"

    return {
        "order_number": order_number,
        "tracking_numbers": tracking_numbers,
        "ship_to": ship_to,
        "is_cancellation": is_cancel,
        "cancellation_reason": cancellation_reason,
        "retailer": retailer
    }


# -----------------------
# CSV utilities & status logic
# -----------------------
def load_existing_orders_csv(csv_file):
    rows = []
    index = {}
    if not csv_file or not os.path.exists(csv_file):
        log(f"Orders CSV {csv_file} not found (will be created if needed).")
        return rows, index
    try:
        with open(csv_file, "r", newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                rows.append(row)
                if row.get("order_number"):
                    index[row["order_number"]] = row
        log(f"Loaded {len(rows)} rows from orders CSV {csv_file}")
    except Exception as e:
        log(f"Could not load orders CSV {csv_file}: {e}")
    return rows, index


def rewrite_orders_csv(rows, csv_file):
    if not csv_file:
        raise ValueError("csv_file required")
    os.makedirs(os.path.dirname(csv_file), exist_ok=True)
    fieldnames = ["order_number", "tracking_numbers", "ship_to", "sent_to", "sent_date", "status", "retailer"]
    try:
        with open(csv_file, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for r in rows:
                write_row = {k: (r.get(k) or "") for k in fieldnames}
                writer.writerow(write_row)
        log(f"Rewrote orders CSV {csv_file} with {len(rows)} rows")
    except Exception as e:
        log(f"Could not rewrite orders CSV {csv_file}: {e}")


def load_existing_csv_orders(csv_file):
    existing = set()
    if not csv_file or not os.path.exists(csv_file):
        return existing
    try:
        with open(csv_file, "r", newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                num = row.get("order_number")
                if num:
                    existing.add(num)
    except Exception as e:
        log(f"Could not read CSV {csv_file}: {e}")
    return existing


def detect_highest_status_for_order(order_entry_list, retailer_hint=None):
    rank = {"ordered": 1, "shipped": 2, "delivered": 3, "cancelled": 4}

    rules = RETAILER_RULES.get(retailer_hint) if retailer_hint else None
    delivered_keywords = rules.get("delivered_indicators", []) if rules else ["delivered", "out for delivery", "arrived"]
    shipped_keywords = rules.get("shipped_indicators", []) if rules else ["shipped", "on the way"]

    # cancellation first
    for entry in order_entry_list:
        ex = entry.get("extracted", {}) or {}
        if ex.get("is_cancellation"):
            return "cancelled"

    # delivered detection
    for entry in order_entry_list:
        html = (entry.get("html") or "").lower()
        text = (BeautifulSoup(html, "lxml").get_text(separator="\n") if html else "").lower()
        for kw in delivered_keywords:
            if kw in html or kw in text:
                return "delivered"

    # shipped detection
    for entry in order_entry_list:
        subj = (entry.get("subject") or "").lower()
        html = (entry.get("html") or "").lower()
        for kw in shipped_keywords:
            if kw in subj or kw in html:
                return "shipped"
        ex = entry.get("extracted", {}) or {}
        if ex.get("tracking_numbers"):
            return "shipped"

    return "ordered"


# -----------------------
# Main processing function (updates and writes)
# -----------------------
def process_and_write_orders(all_cache_data, orders_csv, cancellations_csv):
    grouped = defaultdict(list)
    for entry in all_cache_data:
        extracted = entry.get("extracted", {}) or {}
        order_num = extracted.get("order_number")
        if order_num:
            grouped[order_num].append(entry)

    orders_rows, orders_index = load_existing_orders_csv(orders_csv)
    existing_cancellations = load_existing_csv_orders(cancellations_csv)

    modifications = 0
    cancellations_written = 0
    new_orders_written = 0

    rank = {"ordered": 1, "shipped": 2, "delivered": 3, "cancelled": 4}

    for order_num, entries in grouped.items():
        retailer_hint = None
        for e in entries:
            ex = e.get("extracted", {}) or {}
            if ex.get("retailer"):
                retailer_hint = ex.get("retailer")
                break

        observed_status = detect_highest_status_for_order(entries, retailer_hint=retailer_hint)
        log(f"Order {order_num}: observed status -> {observed_status} (retailer={retailer_hint})")

        # latest date
        latest_date = ""
        try:
            dates = [entry.get("date") for entry in entries if entry.get("date")]
            parsed_dates = []
            for d in dates:
                try:
                    parsed_dates.append(datetime.fromisoformat(d))
                except Exception:
                    try:
                        parsed_dates.append(datetime.fromisoformat(d.replace("Z", "+00:00")))
                    except Exception:
                        pass
            if parsed_dates:
                latest_date = max(parsed_dates).isoformat()
        except Exception:
            latest_date = ""

        # collect info
        tracking = []
        sent_to = ""
        ship_to = ""
        retailer_name = retailer_hint or ""
        for e in entries:
            ex = e.get("extracted", {}) or {}
            tracking.extend(ex.get("tracking_numbers", []) or [])
            if not sent_to and e.get("to"):
                sent_to = e.get("to")
            if not ship_to and ex.get("ship_to"):
                ship_to = ex.get("ship_to")
            if not retailer_name and ex.get("retailer"):
                retailer_name = ex.get("retailer")

        tracking_str = ", ".join(sorted(set(tracking)))

        if order_num in orders_index:
            existing_row = orders_index[order_num]
            current_status = (existing_row.get("status") or "ordered").lower()
            if current_status == "delivered":
                log(f"Order {order_num} already delivered (terminal). Skipping.")
                continue

            if rank.get(observed_status, 0) > rank.get(current_status, 0):
                old_status = current_status
                existing_row["status"] = observed_status.upper() if observed_status != "cancelled" else "CANCELLED"
                if tracking_str:
                    existing_row["tracking_numbers"] = tracking_str
                if ship_to:
                    existing_row["ship_to"] = ship_to
                if sent_to:
                    existing_row["sent_to"] = sent_to
                if latest_date:
                    existing_row["sent_date"] = latest_date
                if retailer_name:
                    existing_row["retailer"] = retailer_name

                modifications += 1
                log(f"Updated order {order_num}: {old_status} -> {existing_row['status']}")

                if observed_status == "cancelled" and order_num not in existing_cancellations:
                    cancel_reason = ""
                    for e in entries:
                        ex = e.get("extracted", {}) or {}
                        if ex.get("is_cancellation"):
                            cancel_reason = ex.get("cancellation_reason", "Not specified")
                            break
                    write_to_csv({
                        "order_number": order_num,
                        "sent_to": sent_to,
                        "sent_date": latest_date,
                        "reason": cancel_reason,
                        "retailer": retailer_name
                    }, cancellations_csv)
                    cancellations_written += 1
                    existing_cancellations.add(order_num)
                    log(f"Wrote cancellation for {order_num} to {cancellations_csv}")

                continue
            else:
                log(f"No advancement for order {order_num} ({current_status} >= {observed_status}).")
                continue
        else:
            # New order
            if observed_status == "cancelled":
                if order_num not in existing_cancellations:
                    cancel_reason = ""
                    for e in entries:
                        ex = e.get("extracted", {}) or {}
                        if ex.get("is_cancellation"):
                            cancel_reason = ex.get("cancellation_reason", "Not specified")
                            break
                    write_to_csv({
                        "order_number": order_num,
                        "sent_to": sent_to,
                        "sent_date": latest_date,
                        "reason": cancel_reason,
                        "retailer": retailer_name
                    }, cancellations_csv)
                    cancellations_written += 1
                    existing_cancellations.add(order_num)
                    log(f"New cancellation {order_num} written to {cancellations_csv}")
            else:
                new_row = {
                    "order_number": order_num,
                    "tracking_numbers": tracking_str,
                    "ship_to": ship_to,
                    "sent_to": sent_to,
                    "sent_date": latest_date,
                    "status": observed_status.upper(),
                    "retailer": retailer_name
                }
                orders_rows.append(new_row)
                orders_index[order_num] = new_row
                new_orders_written += 1
                log(f"Appended new order {order_num} with status {observed_status.upper()} (retailer={retailer_name})")
                continue

    rewrite_orders_csv(orders_rows, orders_csv)

    print(f"\n📊 PROCESS SUMMARY:")
    print(f"   Orders CSV updated rows: {modifications}")
    print(f"   New orders added: {new_orders_written}")
    print(f"   New cancellations written: {cancellations_written}")

    return grouped


# -----------------------
# Main scraping function
# -----------------------
def scrape_target_emails(days_back=DAYS_BACK, email_account=None, password=None, imap_server=None):
    if not email_account or not password or not imap_server:
        raise ValueError("email_account, password and imap_server are required")

    cache_path = CACHE_JSON.replace(".json", f"_{email_account.replace('@', '_')}.json")
    email_cache = load_email_cache(cache_path)
    cached_uids = {entry.get("uid") for entry in email_cache if entry.get("uid")}
    log(f"Cached UIDs for {email_account}: {len(cached_uids)}")

    try:
        log(f"Connecting to {email_account} on {imap_server}...")
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(email_account, password)
        mail.select("inbox")
        log("Connected.")

        typ, data = mail.search(None, "ALL")
        if typ != "OK" or not data or not data[0]:
            log("No messages found.")
            mail.logout()
            return

        cutoff = datetime.now(timezone.utc) - timedelta(days=int(days_back))
        new_cache_entries = []
        target_emails_found = 0

        orders_csv = CSV_PATH.replace('.csv', '_orders.csv') if CSV_PATH else os.path.join(os.getcwd(), "report_orders.csv")
        cancellations_csv = CSV_PATH.replace('.csv', '_cancellations.csv') if CSV_PATH else os.path.join(os.getcwd(), "report_cancellations.csv")

        for num in reversed(data[0].split()):
            uid = num.decode()
            if uid in cached_uids:
                log(f"Skipping UID {uid} (already in cache).")
                continue

            typ, msg_data = mail.fetch(num, "(RFC822)")
            if typ != "OK" or not msg_data:
                log(f"Failed to fetch UID {uid} (server returned {typ}).")
                continue

            msg = email.message_from_bytes(msg_data[0][1])
            subject = clean_subject(msg.get("Subject", ""))
            sender = (msg.get("From") or "").lower()
            to_addr = msg.get("To", "")
            date_hdr = msg.get("Date")

            try:
                sent_dt = parsedate_to_datetime(date_hdr) if date_hdr else None
                sent_dt_utc = sent_dt.astimezone(timezone.utc) if sent_dt else None
            except Exception:
                log(f"UID {uid} - could not parse Date header '{date_hdr}'. Skipping.")
                continue

            if sent_dt_utc and sent_dt_utc < cutoff:
                log(f"UID {uid} from {sent_dt_utc.date()} is older than cutoff {cutoff.date()}, stopping iteration.")
                break

            # Identify retailer (prefer sender then subject)
            retailer = identify_retailer(sender, subject) or identify_retailer(subject, sender) or "unknown"
            if retailer == "unknown":
                if "target" not in sender and "walmart" not in sender and "amazon" not in sender and "bestbuy" not in sender:
                    log(f"Skipping UID {uid} - sender not a known retailer: {sender}")
                    continue
                else:
                    # try a simple substring fallback
                    if "target" in sender:
                        retailer = "target"
                    elif "walmart" in sender:
                        retailer = "walmart"
                    elif "amazon" in sender:
                        retailer = "amazon"
                    elif "bestbuy" in sender or "best buy" in sender:
                        retailer = "bestbuy"

            if retailer not in RETAILER_RULES:
                log(f"Retailer '{retailer}' not supported by rules; skipping UID {uid}.")
                continue

            target_emails_found += 1
            log(f"Processing email #{target_emails_found} (UID {uid}) - Retailer: {retailer} - Subject: {subject}")

            # Extract body
            body = None
            if msg.is_multipart():
                for part in msg.walk():
                    ctype = part.get_content_type()
                    disp = part.get("Content-Disposition") or ""
                    if ctype == "text/html" and "attachment" not in disp:
                        try:
                            body = part.get_payload(decode=True).decode(errors="ignore")
                        except Exception:
                            body = None
                        break
                if not body:
                    for part in msg.walk():
                        if part.get_content_type() == "text/plain" and "attachment" not in (part.get("Content-Disposition") or ""):
                            try:
                                body = part.get_payload(decode=True).decode(errors="ignore")
                            except Exception:
                                body = None
                            break
            else:
                try:
                    body = msg.get_payload(decode=True).decode(errors="ignore")
                except Exception:
                    body = None

            if not body:
                log(f"UID {uid} - no body found, skipping.")
                continue

            soup = BeautifulSoup(body, "lxml")
            text = soup.get_text(separator="\n")

            extracted = extract_order_details(text, subject)
            extracted.setdefault("retailer", retailer)

            cache_entry = {
                "uid": uid,
                "subject": subject,
                "from": sender,
                "to": to_addr,
                "date": sent_dt_utc.isoformat() if sent_dt_utc else "",
                "html": body,
                "extracted": extracted,
            }

            new_cache_entries.append(cache_entry)
            cached_uids.add(uid)

            log(f"UID {uid} added to new cache (order={extracted.get('order_number')}, "
                f"cancel={extracted.get('is_cancellation')}, tracks={len(extracted.get('tracking_numbers', []))}, retailer={extracted.get('retailer')})")

        # Save merged cache (per-account)
        all_cache_data = email_cache + new_cache_entries
        if new_cache_entries:
            save_email_cache(all_cache_data, cache_path)
        else:
            log("No new retailer messages found; cache unchanged.")

        # Process and update CSVs
        process_and_write_orders(all_cache_data, orders_csv, cancellations_csv)

        mail.logout()
        log(f"Finished processing account {email_account}.")

    except Exception as e:
        log(f"Error while scraping {email_account}: {e}")
