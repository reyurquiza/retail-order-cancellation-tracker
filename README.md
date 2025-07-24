# 🎯 Target Email Scraper

This tool connects to your Gmail inbox, scans for recent shipping confirmation emails (like those from Target), and extracts:

- ✅ Order number  
- ✅ Tracking numbers (UPS, FedEx, USPS)  
- ✅ Ship-to address  
- ✅ Recipient email and sent date  

Results are saved to a CSV file and all emails (with or without data) are cached to avoid duplicate processing.

---

## 🚀 Features

- 🔒 Secure IMAP login using Gmail app password  
- ⏱ Only scans emails from the past X days  
- 📨 Caches processed emails to avoid re-checking  
- 📁 Outputs clean CSV reports  
- 📬 Extracts from HTML emails using BeautifulSoup

---

## 📦 Setup

### 1. Clone the repository

```bash
git clone https://github.com/zzpixels/targetemailscraper.git
cd targetemailscraper
```

### 2. Install dependencies

We recommend using a virtual environment:

```bash
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install -r requirements.txt
```

### 3. Create `config.py`

Make a new file in the project root called `config.py`: (Or use the one provided)

```python
EMAIL = "your_email@gmail.com"
PASSWORD = "your_app_password"  # Gmail app password, not your regular password
IMAP_SERVER = "imap.gmail.com"
DAYS_BACK = 7  # How many days back to look for emails
CSV_PATH = "output/report.csv"
CACHE_JSON = "cache/emails.json"
```

> 🔐 **Important**: You need to [generate an App Password in Gmail](https://support.google.com/mail/answer/185833) if you use 2-Step Verification.

---

## ▶️ Usage

Run the scraper:

```bash
python script.py
```

This will:
- Connect to your Gmail inbox
- Scan emails from the last `DAYS_BACK` days
- Extract shipment data (order number, tracking numbers, address)
- Save results to `output/report.csv`
- Cache scanned emails in `cache/emails.json`

---

## 🧪 Example Output

CSV: `output/report.csv`

| order_number     | tracking_numbers        | ship_to                 | sent_to               | sent_date             |
|------------------|-------------------------|--------------------------|------------------------|------------------------|
| 902002669367042  | 1Z999AA10123456784      | 123 Main St, NY 10001    | abc@example.com        | 2025-07-21T17:50:54Z   |

---

## 📁 Folder Structure

```
target-email-scraper/
├── scraper.py           # Main script
├── config.py            # Your email settings
├── requirements.txt     # Python dependencies
├── output/
│   └── report.csv       # Generated report
├── cache/
│   └── emails.json      # Cached parsed emails
```

---

## 📄 License

MIT — Feel free to use, modify, and share this tool.
