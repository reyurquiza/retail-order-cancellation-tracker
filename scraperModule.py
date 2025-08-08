"""
emailScraper.py
---------------
Streamlined backend module for scraping Target order and cancellation emails,
with verbose logging so the GUI can display every backend action.

Behavior:
- Connect via IMAP and parse Target emails.
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
# Helpers / I/O
# -----------------------
def log(msg):
    """Console logging helper (captured by GUI's RealTimeLogger)."""
    print(f"[LOG] {msg}")


def load_email_cache(filename):
    """Load cache from the given filename, return list (empty if none)."""
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
    """Save cache_data to the given filename (creates directories as needed)."""
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
    """
    Append a row dict to csv_file. Creates parent dir and headers if needed.
    Chooses schema based on filename suffix ('_orders.csv' uses the order schema).
    """
    if not csv_file:
        raise ValueError("csv_file path required")
    os.makedirs(os.path.dirname(csv_file), exist_ok=True)
    file_needs_header = not os.path.isfile(csv_file) or os.stat(csv_file).st_size == 0
    mode = "a"
    with open(csv_file, mode, newline="", encoding="utf-8") as f:
        if csv_file.endswith("_orders.csv"):
            fieldnames = ["order_number", "tracking_numbers", "ship_to", "sent_to", "sent_date", "status"]
        else:
            fieldnames = ["order_number", "sent_to", "sent_date", "reason"]
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        if file_needs_header:
            writer.writeheader()
            log(f"Wrote header to {csv_file}")
        writer.writerow(row)
        log(f"Wrote row to {csv_file}: {row.get('order_number')}")


# -----------------------
# Parsing helpers
# -----------------------
def clean_subject(subject):
    """Decode email subject (handles encoded parts)."""
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


def is_cancellation_email(subject, text):
    """Return True if subject/body indicate a cancellation notice."""
    s = (subject or "").lower()
    t = (text or "").lower()
    checks = [
        "sorry, we had to cancel" in s,
        "cancel order" in s,
        "canceled" in s,
        "cancelled" in s,
        "your order has been canceled" in t,
        "your order has been cancelled" in t,
        "order was canceled" in t,
        "order was cancelled" in t,
        "we had to cancel" in t,
        "purchase limit exceeded" in t,
        "payment issue" in t,
        "activity not supported" in t,
        "you haven't been charged" in t,
        "system automatically canceled" in t,
    ]
    return any(checks)


def extract_order_details(text, subject):
    """
    Parse plain-text email content and subject to extract:
      - order_number (8-15 digits expected)
      - tracking_numbers (FedEx/UPS/USPS heuristics)
      - ship_to (address if present)
      - is_cancellation (bool)
      - cancellation_reason (if cancellation)
    """
    order_patterns = [
        r"order\s*#?\s*(\d{8,15})",
        r"order\s*number[:\s]*(\d{8,15})",
        r"#(\d{8,15})",
    ]

    order_number = None
    for pattern in order_patterns:
        m = re.search(pattern, text or "", re.IGNORECASE) or re.search(pattern, subject or "", re.IGNORECASE)
        if m:
            order_number = m.group(1)
            break

    cancellation_flag = is_cancellation_email(subject or "", text or "")

    tracking_numbers = []
    if not cancellation_flag:
        fedex = re.findall(r"\b(\d{12,22})\b", text or "")
        ups = re.findall(r"\b(1Z[0-9A-Z]{16})\b", text or "")
        usps = re.findall(r"\b(\d{20,22}|\d{13})\b", text or "")
        tracking_numbers = list(set(fedex + ups + usps))
        if order_number in tracking_numbers:
            tracking_numbers.remove(order_number)

    ship_to = None
    m_addr = re.search(r"Delivers to:\s*(.+,\s*\d{5})", text or "", re.IGNORECASE)
    if m_addr:
        ship_to = m_addr.group(1).strip()

    cancellation_reason = None
    if cancellation_flag:
        t = (text or "").lower()
        if "purchase limit exceeded" in t:
            cancellation_reason = "Purchase limit exceeded"
        elif "payment issue" in t:
            cancellation_reason = "Payment issue"
        elif "activity not supported" in t:
            cancellation_reason = "Activity not supported on Target.com"
        elif "out of stock" in t:
            cancellation_reason = "Out of stock"
        else:
            reason_patterns = [
                r"what went wrong\?\s*(.+?)(?:\*\*|$)",
                r"reason[:\s]*([^.\n]+)",
                r"because[:\s]*([^.\n]+)",
                r"unfortunately[:\s]*([^.\n]+)",
            ]
            for pat in reason_patterns:
                mm = re.search(pat, text or "", re.IGNORECASE | re.DOTALL)
                if mm:
                    cancellation_reason = mm.group(1).strip()[:200]
                    break
            if not cancellation_reason:
                cancellation_reason = "Reason not specified"

    return {
        "order_number": order_number,
        "tracking_numbers": tracking_numbers,
        "ship_to": ship_to,
        "is_cancellation": cancellation_flag,
        "cancellation_reason": cancellation_reason,
    }


# -----------------------
# CSV utilities & status logic
# -----------------------
def load_existing_orders_csv(csv_file):
    """
    Load orders CSV entirely into a list of dicts and return (rows_list, index_map).
    index_map maps order_number -> row_dict (the same object from rows_list).
    If file missing, returns ([], {}).
    """
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
    """Overwrite the orders csv with provided rows. Field order preserved."""
    if not csv_file:
        raise ValueError("csv_file required")
    os.makedirs(os.path.dirname(csv_file), exist_ok=True)
    fieldnames = ["order_number", "tracking_numbers", "ship_to", "sent_to", "sent_date", "status"]
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
    """Return set of order_number values already present in csv_file (empty if missing)."""
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


def detect_highest_status_for_order(order_entry_list):
    """
    Given all cached email entries for a single order (list of email entries),
    detect the highest status observed:
      - 'cancelled' if any email is cancellation
      - 'delivered' if any email contains delivered/out-for-delivery indicators
      - 'shipped' if any tracking number appears
      - else 'ordered'
    """
    status = "ordered"
    delivered_keywords = [
        "delivered",
        "out for delivery",
        "delivery completed",
        "left at the",
        "was delivered",
        "arrived at",
        "delivered on",
        "signature required"
    ]

    # Cancellation overrides everything
    for entry in order_entry_list:
        extracted = entry.get("extracted", {}) or {}
        if extracted.get("is_cancellation"):
            return "cancelled"
    # Delivered detection
    for entry in order_entry_list:
        html = (entry.get("html") or "").lower()
        if any(kw in html for kw in delivered_keywords):
            return "delivered"
    # Shipped detection
    for entry in order_entry_list:
        extracted = entry.get("extracted", {}) or {}
        if extracted.get("tracking_numbers"):
            return "shipped"
    return "ordered"


# -----------------------
# Main processing function (updates and writes)
# -----------------------
def process_and_write_orders(all_cache_data, orders_csv, cancellations_csv):
    """
    New behavior:
    - Group all cached messages by order_number.
    - For each order:
        * compute highest status observed from cached messages
        * if the order exists in orders CSV and observed status is an advancement,
          update the CSV (and stop further processing for that order).
        * if cancellation observed and not yet in cancellations CSV -> write to cancellations CSV.
        * if order is new -> append to orders CSV.
    - After processing all orders, rewrite the orders CSV to persist updates.
    """
    grouped = defaultdict(list)
    for entry in all_cache_data:
        extracted = entry.get("extracted", {}) or {}
        order_num = extracted.get("order_number")
        if order_num:
            grouped[order_num].append(entry)

    orders_rows, orders_index = load_existing_orders_csv(orders_csv)
    existing_orders_set = set(orders_index.keys())
    existing_cancellations = load_existing_csv_orders(cancellations_csv)

    appended_rows = []
    modifications = 0
    cancellations_written = 0
    new_orders_written = 0

    rank = {"ordered": 1, "shipped": 2, "delivered": 3, "cancelled": 4}

    for order_num, entries in grouped.items():
        observed_status = detect_highest_status_for_order(entries)
        log(f"Order {order_num}: observed status -> {observed_status}")

        # Determine latest date (ISO)
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

        # Gather contact/tracking info
        tracking = []
        sent_to = ""
        ship_to = ""
        for e in entries:
            ex = e.get("extracted", {}) or {}
            tracking.extend(ex.get("tracking_numbers", []) or [])
            if not sent_to and e.get("to"):
                sent_to = e.get("to")
            if not ship_to and ex.get("ship_to"):
                ship_to = ex.get("ship_to")

        tracking_str = ", ".join(sorted(set(tracking)))

        if order_num in orders_index:
            existing_row = orders_index[order_num]
            current_status = (existing_row.get("status") or "ordered").lower()
            log(f"Existing order {order_num} current status: {current_status}")

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
                        "reason": cancel_reason
                    }, cancellations_csv)
                    cancellations_written += 1
                    existing_cancellations.add(order_num)
                    log(f"Wrote cancellation for {order_num} to {cancellations_csv}")

                # Once we've applied an advancement, stop further processing for this order this run
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
                        "reason": cancel_reason
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
                    "status": observed_status.upper()
                }
                orders_rows.append(new_row)
                orders_index[order_num] = new_row
                new_orders_written += 1
                log(f"Appended new order {order_num} with status {observed_status.upper()}")
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
    """
    Connect to an IMAP server for a single account, fetch Target emails within the
    given days_back window, parse and append them to a per-account cache, then
    process the aggregated cache to update CSVs (including status updates).
    """
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

        orders_csv = CSV_PATH.replace(".csv", "_orders.csv")
        cancellations_csv = CSV_PATH.replace(".csv", "_cancellations.csv")

        # Iterate newest -> oldest
        for num in reversed(data[0].split()):
            uid = num.decode()
            if uid in cached_uids:
                # Inform GUI that we're skipping known UIDs
                log(f"Skipping UID {uid} (already in cache).")
                continue  # already processed for this account

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

            # Only process Target sender emails
            if "target" not in sender:
                log(f"Skipping UID {uid} - sender not Target: {sender}")
                continue

            target_emails_found += 1
            log(f"Processing email #{target_emails_found} (UID {uid}) - Subject: {subject}")

            # Extract HTML body
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
            else:
                try:
                    body = msg.get_payload(decode=True).decode(errors="ignore")
                except Exception:
                    body = None

            if not body:
                log(f"UID {uid} - no HTML body found, skipping.")
                continue

            text = BeautifulSoup(body, "lxml").get_text(separator="\n")
            extracted = extract_order_details(text, subject)

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
            cached_uids.add(uid)  # avoid duplicates within same run

            # Inform GUI what we extracted for this email (short summary)
            log(f"UID {uid} added to new cache (order={extracted.get('order_number')}, "
                f"cancel={extracted.get('is_cancellation')}, tracks={len(extracted.get('tracking_numbers', []))})")

        # Merge and persist per-account cache
        all_cache_data = email_cache + new_cache_entries
        if new_cache_entries:
            save_email_cache(all_cache_data, cache_path)
        else:
            log("No new Target messages found; cache unchanged.")

        # Process and write CSVs using aggregated cache (this now updates statuses)
        process_and_write_orders(all_cache_data, orders_csv, cancellations_csv)

        mail.logout()
        log(f"Finished processing account {email_account}.")

    except Exception as e:
        log(f"Error while scraping {email_account}: {e}")


# -----------------------
# If run directly, iterate accounts from config
# -----------------------
if __name__ == "__main__":
    print(f"🚀 Starting Target email scraping for {len(EMAIL_ACCOUNTS)} account(s)...")
    for i, account in enumerate(EMAIL_ACCOUNTS, 1):
        print("\n" + "=" * 70)
        print(f"Processing account {i}/{len(EMAIL_ACCOUNTS)}: {account.get('email')}")
        print("=" * 70)
        try:
            scrape_target_emails(
                days_back=DAYS_BACK,
                email_account=account.get("email"),
                password=account.get("password"),
                imap_server=account.get("imap_server"),
            )
        except Exception as e:
            log(f"Error processing {account.get('email')}: {e}")
            continue
    print(f"🎉 Finished processing all {len(EMAIL_ACCOUNTS)} accounts!")
