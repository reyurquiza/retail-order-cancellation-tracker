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


# ─────────── HELPERS ───────────
def log(msg):
    print(f"[LOG] {msg}")


def load_email_cache(filename):
    if os.path.exists(CACHE_JSON):
        with open(CACHE_JSON, "r") as f:
            return json.load(f)
    return []

def save_email_cache(cache_data, filename):
    os.makedirs(os.path.dirname(CACHE_JSON), exist_ok=True)
    with open(CACHE_JSON, "w") as f:
        json.dump(cache_data, f, indent=2)

def write_to_csv(row, csv_file):
    os.makedirs(os.path.dirname(csv_file), exist_ok=True)
    file_needs_header = not os.path.isfile(csv_file) or os.stat(csv_file).st_size == 0
    with open(csv_file, "a+", newline="") as f:
        if csv_file.endswith('_orders.csv'):
            fieldnames = ["order_number", "tracking_numbers", "ship_to", "sent_to", "sent_date", "status"]
        else:
            fieldnames = ["order_number", "sent_to", "sent_date", "reason"]
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        if file_needs_header:
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


def is_cancellation_email(subject, text):
    """Check if this is a cancellation email"""
    # Target-specific cancellation indicators
    subject_lower = subject.lower()
    text_lower = text.lower()

    # Check for Target's specific cancellation patterns
    cancellation_indicators = [
        # Subject line patterns
        'sorry, we had to cancel' in subject_lower,
        'cancel order' in subject_lower,
        'canceled' in subject_lower,
        'cancelled' in subject_lower,

        # Body text patterns
        'your order has been canceled' in text_lower,
        'your order has been cancelled' in text_lower,
        'order was canceled' in text_lower,
        'order was cancelled' in text_lower,
        'we had to cancel' in text_lower,
        'purchase limit exceeded' in text_lower,
        'payment issue' in text_lower,
        'activity not supported' in text_lower,
        'you haven\'t been charged' in text_lower,
        'system automatically canceled' in text_lower
    ]

    return any(cancellation_indicators)


def extract_order_details(text, subject):
    """Extract order number and other details"""
    # Target order numbers can be 8-15 digits
    order_patterns = [
        r'order\s*#?\s*(\d{8,15})',
        r'order\s*number[:\s]*(\d{8,15})',
        r'#(\d{8,15})'
    ]

    order_number = None
    for pattern in order_patterns:
        order_match = re.search(pattern, text, re.IGNORECASE)
        if not order_match:
            order_match = re.search(pattern, subject, re.IGNORECASE)
        if order_match:
            order_number = order_match.group(1)
            break

    # Extract tracking numbers (only for non-cancellation emails)
    tracking_numbers = []
    if not is_cancellation_email(subject, text):
        fedex_match = re.findall(r'\b(\d{12,22})\b', text)
        ups_match = re.findall(r'\b(1Z[0-9A-Z]{16})\b', text)
        usps_match = re.findall(r'\b(\d{20,22}|\d{13})\b', text)

        tracking_numbers = list(set(fedex_match + ups_match + usps_match))
        if order_number and order_number in tracking_numbers:
            tracking_numbers.remove(order_number)

    # Extract shipping address
    address_match = re.search(r'Delivers to:\s*(.+,\s*\d{5})', text)
    ship_to = address_match.group(1) if address_match else None

    # Extract cancellation reason
    cancellation_reason = None
    if is_cancellation_email(subject, text):
        # Target-specific cancellation reasons
        if 'purchase limit exceeded' in text.lower():
            cancellation_reason = "Purchase limit exceeded"
        elif 'payment issue' in text.lower():
            cancellation_reason = "Payment issue"
        elif 'activity not supported' in text.lower():
            cancellation_reason = "Activity not supported on Target.com"
        elif 'out of stock' in text.lower():
            cancellation_reason = "Out of stock"
        else:
            # Generic reason extraction
            reason_patterns = [
                r'what went wrong\?\s*(.+?)(?:\*\*|$)',
                r'reason[:\s]*([^.\n]+)',
                r'because[:\s]*([^.\n]+)',
                r'unfortunately[:\s]*([^.\n]+)'
            ]
            for pattern in reason_patterns:
                reason_match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
                if reason_match:
                    cancellation_reason = reason_match.group(1).strip()[:100]
                    break

            if not cancellation_reason:
                cancellation_reason = "Reason not specified"

    return {
        "order_number": order_number,
        "tracking_numbers": tracking_numbers,
        "ship_to": ship_to,
        "is_cancellation": is_cancellation_email(subject, text),
        "cancellation_reason": cancellation_reason
    }


def load_existing_csv_orders(csv_file):
    existing_orders = set()
    if os.path.exists(csv_file):
        try:
            with open(csv_file, 'r', newline='') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    existing_orders.add(row['order_number'])
            # log(f"Loaded {len(existing_orders)} existing order numbers from {csv_file}")
        except Exception as e:
            log(f"⚠️ Warning: Could not read CSV {csv_file}: {e}")
    else:
        log(f"ℹ️ CSV file {csv_file} does not exist yet.")
    return existing_orders



def process_and_write_orders(all_cache_data, orders_csv, cancellations_csv):
    """Process all emails and write to appropriate CSV files based on final order status"""
    orders = defaultdict(lambda: {
        'status': 'unknown',
        'order_emails': [],
        'cancel_emails': [],
        'tracking_numbers': [],
        'ship_to': None,
        'sent_to': None
    })

    # Load existing orders to avoid duplicates
    existing_orders = load_existing_csv_orders(orders_csv)
    existing_cancellations = load_existing_csv_orders(cancellations_csv)

    # First pass: collect all emails by order number
    for entry in all_cache_data:
        extracted = entry.get('extracted', {})
        order_num = extracted.get('order_number')

        if not order_num:
            continue

        if extracted.get('is_cancellation'):
            orders[order_num]['cancel_emails'].append(entry)
            orders[order_num]['status'] = 'cancelled'  # Mark as cancelled
        else:
            orders[order_num]['order_emails'].append(entry)
            if orders[order_num]['status'] != 'cancelled':  # Don't override cancellation
                orders[order_num]['status'] = 'ordered'

            # Collect tracking info and other details
            if extracted.get('tracking_numbers'):
                orders[order_num]['tracking_numbers'].extend(extracted['tracking_numbers'])
            if extracted.get('ship_to'):
                orders[order_num]['ship_to'] = extracted['ship_to']
            if entry.get('to'):
                orders[order_num]['sent_to'] = entry['to']

    # Second pass: write to appropriate CSV based on final status (only new orders)
    orders_written = 0
    cancellations_written = 0

    for order_num, order_data in orders.items():
        if order_data['status'] == 'cancelled':
            # Write to cancellations CSV (only if not already there)
            if order_num not in existing_cancellations:
                cancel_email = order_data['cancel_emails'][0] if order_data['cancel_emails'] else None
                if cancel_email:
                    cancel_reason = cancel_email.get('extracted', {}).get('cancellation_reason', 'Not specified')
                    cancel_date = cancel_email.get('date', '')

                    write_to_csv({
                        "order_number": order_num,
                        "sent_to": order_data['sent_to'] or cancel_email.get('to', ''),
                        "sent_date": cancel_date,
                        "reason": cancel_reason
                    }, cancellations_csv)
                    cancellations_written += 1
                    print(f"❌ Order {order_num} - CANCELLED (new)")
            else:
                print(f"❌ Order {order_num} - CANCELLED (already in CSV)")
        else:
            # Write to orders CSV (non-cancelled orders only, and only if not already there)
            if order_num not in existing_orders:
                order_email = order_data['order_emails'][0] if order_data['order_emails'] else None
                if order_email:
                    # Remove duplicates from tracking numbers
                    unique_tracking = list(set(order_data['tracking_numbers']))

                    write_to_csv({
                        "order_number": order_num,
                        "tracking_numbers": ", ".join(unique_tracking),
                        "ship_to": order_data['ship_to'],
                        "sent_to": order_data['sent_to'] or order_email.get('to', ''),
                        "sent_date": order_email.get('date', ''),
                        "status": "shipped" if unique_tracking else "ordered"
                    }, orders_csv)
                    orders_written += 1
                    if unique_tracking:
                        print(f"✅ Order {order_num} - SHIPPED (new)")
                    else:
                        print(f"📦 Order {order_num} - ORDERED (new)")
            else:
                print(f"📦 Order {order_num} - Already in orders CSV")

    print(f"\n📊 NEW ENTRIES ADDED:")
    print(f"   New orders written to CSV: {orders_written}")
    print(f"   New cancellations written to CSV: {cancellations_written}")

    return orders


# ─────────── MAIN ───────────
def scrape_target_emails(days_back=DAYS_BACK, email_account=None, password=None, imap_server=None):
    try:
        # Use provided credentials or fall back to config
        email_to_use = email_account or EMAIL
        password_to_use = password or PASSWORD
        imap_to_use = imap_server or IMAP_SERVER

        log(f"⚠️ Connecting to {email_to_use} on {imap_to_use}...")

        cache_path = CACHE_JSON.replace(".json", f"_{email_to_use.replace('@', '_')}.json")
        email_cache = load_email_cache(CACHE_JSON)

        cached_uids = {entry["uid"] for entry in email_cache}  # Use cache to prevent duplicates

        mail = imaplib.IMAP4_SSL(imap_to_use)
        mail.login(email_to_use, password_to_use)
        mail.select("inbox")

        log(f"✅ Connected!")

        typ, data = mail.search(None, "ALL")
        if typ != "OK" or not data or not data[0]:
            log("❌ No messages found.")
            return

        cutoff = datetime.now(timezone.utc) - timedelta(days=days_back)
        new_cache_entries = []
        processed_count = 0
        target_emails_found = 0

        # Separate CSV files for orders and cancellations
        orders_csv = CSV_PATH.replace('.csv', '_orders.csv')
        cancellations_csv = CSV_PATH.replace('.csv', '_cancellations.csv')

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
                log(f"🛑 UID {uid} from {sent_dt_utc.date()} is older than cutoff {cutoff.date()}, stopping.")
                break

            # Only process Target emails
            if 'target' not in sender:
                continue

            target_emails_found += 1
            print(f"\n📧 Processing {sender.split()[0].upper()} email #{target_emails_found}...")
            print(f"   Subject: {subject}")

            body = None
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/html" and "attachment" not in (
                            part.get("Content-Disposition") or ""):
                        body = part.get_payload(decode=True).decode(errors="ignore")
                        break
            else:
                body = msg.get_payload(decode=True).decode(errors="ignore")

            if not body:
                continue

            soup = BeautifulSoup(body, "lxml")
            text = soup.get_text(separator="\n")

            extracted = extract_order_details(text, subject)

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

        # Save combined cache
        all_cache_data = email_cache + new_cache_entries
        #save_email_cache(all_cache_data, CACHE_JSON)
        log(f"✅ Done. {len(new_cache_entries)} new messages processed.")

        # Process all orders and write to CSV files
        orders = process_and_write_orders(all_cache_data, orders_csv, cancellations_csv)

        # Generate summary report
        total_orders = len(orders)
        cancelled_orders = sum(1 for order in orders.values() if order['status'] == 'cancelled')
        successful_orders = total_orders - cancelled_orders

        print(f"\n📊 ORDER SUMMARY REPORT:")
        print(f"   Total unique orders: {total_orders}")
        print(f"   Cancelled orders: {cancelled_orders}")
        print(f"   Successful orders: {successful_orders}")
        if total_orders > 0:
            print(f"   Cancellation rate: {cancelled_orders / total_orders:.1%}")

        mail.logout()

    except Exception as e:
        log(f"❌ Error: {e}")


# ─────────── RUN ───────────
if __name__ == "__main__":
    print(f"🚀 Starting Target email scraping for {len(EMAIL_ACCOUNTS)} account(s)...")

    for i, account in enumerate(EMAIL_ACCOUNTS, 1):
        print(f"\n{'=' * 70}")
        print(f"Processing account {i}/{len(EMAIL_ACCOUNTS)}: {account['email']}")
        print(f"{'=' * 70}")

        try:
            scrape_target_emails(
                email_account=account['email'],
                password=account['password'],
                imap_server=account['imap_server']
            )
        except Exception as e:
            log(f"❌ Error processing {account['email']}: {e}")
            continue

    print(f"🎉 Finished processing all {len(EMAIL_ACCOUNTS)} accounts!")