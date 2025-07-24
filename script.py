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

from config import EMAIL, PASSWORD, IMAP_SERVER, DAYS_BACK, CSV_PATH, CACHE_JSON

# ─────────── HELPERS ───────────
def log(msg):
    print(f"[LOG] {msg}")


def load_email_cache():
    if os.path.exists(CACHE_JSON):
        with open(CACHE_JSON, "r") as f:
            return json.load(f)
    return []


def save_email_cache(cache_data):
    os.makedirs(os.path.dirname(CACHE_JSON), exist_ok=True)
    with open(CACHE_JSON, "w") as f:
        json.dump(cache_data, f, indent=2)


def write_to_csv(row):
    os.makedirs(os.path.dirname(CSV_PATH), exist_ok=True)
    file_exists = os.path.isfile(CSV_PATH)
    with open(CSV_PATH, "a", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["order_number", "tracking_numbers", "ship_to", "sent_to", "sent_date"])
        if not file_exists:
            writer.writeheader()
        writer.writerow(row)


def clean_subject(subject):
    decoded_parts = decode_header(subject)
    subject_str = ''
    for part, encoding in decoded_parts:
        if isinstance(part, bytes):
            subject_str += part.decode(encoding or 'utf-8', errors='ignore')
        else:
            subject_str += part
    return subject_str


def extract_details(text):
    order_match = re.search(r'order\s+#?(\d{15})', text, re.IGNORECASE)
    order_number = order_match.group(1) if order_match else None

    fedex_match = re.findall(r'\b(\d{12,22})\b', text)
    ups_match = re.findall(r'\b(1Z[0-9A-Z]{16})\b', text)
    usps_match = re.findall(r'\b(\d{20,22}|\d{13})\b', text)

    tracking_numbers = list(set(fedex_match + ups_match + usps_match))
    if order_number in tracking_numbers:
        tracking_numbers.remove(order_number)

    address_match = re.search(r'Delivers to:\s*(.+,\s*\d{5})', text)

    return {
        "order_number": order_number,
        "tracking_numbers": tracking_numbers,
        "ship_to": address_match.group(1) if address_match else None
    }


# ─────────── MAIN ───────────
def scrape_target_emails(days_back=DAYS_BACK):
    try:
        email_cache = load_email_cache()
        cached_uids = {entry["uid"] for entry in email_cache}

        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL, PASSWORD)
        mail.select("inbox")

        typ, data = mail.search(None, "ALL")
        if typ != "OK" or not data or not data[0]:
            log("❌ No messages found.")
            return

        cutoff = datetime.now(timezone.utc) - timedelta(days=days_back)
        new_cache_entries = []

        for num in reversed(data[0].split()):
            uid = num.decode()
            if uid in cached_uids:
                continue

            typ, msg_data = mail.fetch(num, '(RFC822)')
            if typ != "OK":
                continue

            msg = email.message_from_bytes(msg_data[0][1])
            subject = clean_subject(msg.get("Subject", ""))
            sender = msg.get("From", "").lower()
            to_addr = msg.get("To", "")
            date_hdr = msg.get("Date")

            try:
                sent_dt = parsedate_to_datetime(date_hdr) if date_hdr else None
                sent_dt_utc = sent_dt.astimezone(timezone.utc)
            except:
                continue

            if sent_dt_utc < cutoff:
                log(f"🛑 UID {uid} too old, stopping.")
                break

            body = None
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/html" and "attachment" not in (part.get("Content-Disposition") or ""):
                        body = part.get_payload(decode=True).decode(errors="ignore")
                        break
            else:
                body = msg.get_payload(decode=True).decode(errors="ignore")

            if not body:
                continue

            soup = BeautifulSoup(body, "lxml")
            text = soup.get_text(separator="\n")

            extracted = extract_details(text)

            cache_entry = {
                "uid": uid,
                "subject": subject,
                "from": sender,
                "to": to_addr,
                "date": sent_dt_utc.isoformat(),
                "html": body,
                "extracted": extracted
            }

            new_cache_entries.append(cache_entry)

            if extracted["order_number"] and extracted["tracking_numbers"]:
                write_to_csv({
                    "order_number": extracted["order_number"],
                    "tracking_numbers": ", ".join(extracted["tracking_numbers"]),
                    "ship_to": extracted["ship_to"],
                    "sent_to": to_addr,
                    "sent_date": sent_dt_utc.strftime("%Y-%m-%dT%H:%M:%SZ")
                })
                print(f"✅ UID {uid} exported to CSV")

        # Save combined cache
        save_email_cache(email_cache + new_cache_entries)
        log(f"✅ Done. {len(new_cache_entries)} new messages cached.")

        mail.logout()

    except Exception as e:
        log(f"❌ Error: {e}")


# ─────────── RUN ───────────
if __name__ == "__main__":
    scrape_target_emails()
