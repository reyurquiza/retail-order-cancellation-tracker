# Retail Order Email Scraper

A powerful Python application that connects to your email accounts and extracts order information from major retailers including Target, Walmart, Amazon, and Best Buy. The scraper automatically categorizes orders by status (ordered, shipped, delivered, cancelled) and exports the data to CSV files for easy analysis.

---

## Table of contents

- [Features](#features)  
- [Installation](#installation)  
- [Quick Start](#quick-start)
- [Email Setup](#email-setup)
- [Output Files](#output-files)
- [Adding New Retailers](#adding-new-retailers)
- [GUI Features](#gui-features)
- [Troubleshooting](#troubleshooting)
- [License](#license)

---

## Features

- **Multi-Retailer Support**: Target, Walmart, Amazon, Best Buy (easily extensible)
- **Smart Order Tracking**: Automatically detects order numbers, tracking numbers, and shipping addresses
- **Status Management**: Tracks order progression from ordered → shipped → delivered
- **Cancellation Detection**: Identifies and logs cancelled orders with reasons
- **Real-time GUI**: User-friendly interface with live progress updates
- **CSV Export**: Generates detailed reports in CSV format
- **Email Caching**: Avoids reprocessing emails with intelligent caching
- **Multiple Accounts**: Support for multiple email accounts per retailer

## Installation

### Option 1: Download Pre-built Executable (Recommended for End Users)

1. **Download the latest release:**
   - Go to the [Releases](https://github.com/your-username/target-email-scraper/releases) page
   - Download `TargetEmailScraper.exe` from the latest release
   - No Python installation required!

2. **Run the application:**
   - Double-click `TargetEmailScraper.exe` to launch
   - Windows may show a security warning - click "More info" → "Run anyway"

### Option 2: Run from Source Code (For Developers)

1. **Clone or download the repository**
2. **Install required dependencies:**
   ```bash
   pip install beautifulsoup4 lxml
   ```
3. **Ensure you have Python 3.8+ installed**
4. **Run with:** `python gui.py`

## Quick Start

### Using the Pre-built Executable (Easiest)

1. **Download and run `TargetEmailScraper.exe`**
2. **Configure your email accounts** (see Configuration section below)
3. **Start scraping and view results**

### Using the GUI (Recommended)

1. **Run the GUI application:**
   ```bash
   python gui.py
   ```

2. **Configure your email accounts:**
   - Go to the "Configuration" tab
   - Click "Add Account" to add email accounts
   - Enter your email, password, and IMAP server
   - Set the number of days to scan back (default: 7)
   - Choose an output folder for results

3. **Save configuration and run:**
   - Click "Save Configuration"
   - Switch to the "Run Scraper" tab
   - Click "Start Scraping"
   - Monitor progress in real-time

4. **View results:**
   - Switch to the "View Results" tab
   - Click "Refresh Results" to see the latest data
   - Right-click on orders to copy tracking numbers or hide items
   - Click "Open Output Folder" to access CSV files

### Using the Command Line

1. **Edit the configuration file (`config.py`):**
   ```python
   EMAIL_ACCOUNTS = [
       {
           "email": "your-email@gmail.com",
           "password": "your-app-password",
           "imap_server": "imap.gmail.com"
       }
   ]
   
   DAYS_BACK = 7
   OUTPUT_DIR = "C:/path/to/output"
   CSV_PATH = OUTPUT_DIR + "/report.csv"
   CACHE_JSON = OUTPUT_DIR + "/cache/emails.json"
   ```

2. **Run the scraper programmatically:**
   ```python
   from scraperModule import scrape_target_emails
   
   scrape_target_emails(
       days_back=7,
       email_account="your-email@gmail.com", 
       password="your-app-password",
       imap_server="imap.gmail.com"
   )
   ```

## Email Setup

### Gmail
- **IMAP Server:** `imap.gmail.com`
- **Enable 2FA** and generate an **App Password**
- **Enable IMAP** in Gmail settings

### Outlook/Hotmail
- **IMAP Server:** `imap-mail.outlook.com`
- **Use your regular password** or app-specific password

### Yahoo Mail
- **IMAP Server:** `imap.mail.yahoo.com`
- **Generate an App Password** in Yahoo Account Security

### Other Providers
Check your email provider's IMAP settings documentation.

---

## Output Files

The scraper generates several CSV files in your specified output directory:

- **`report_orders.csv`** - All orders with tracking information
- **`report_cancellations.csv`** - Cancelled orders with reasons
- **`cache/emails_[account].json`** - Cached email data (per account)

### CSV Structure

**Orders CSV:**
- `order_number` - The order/tracking number
- `tracking_numbers` - Shipping tracking numbers
- `ship_to` - Shipping address
- `sent_to` - Email address
- `sent_date` - Email timestamp
- `status` - Order status (ORDERED, SHIPPED, DELIVERED, CANCELLED)
- `retailer` - Store name (target, walmart, amazon, bestbuy)

**Cancellations CSV:**
- `order_number` - The cancelled order number
- `sent_to` - Email address
- `sent_date` - Cancellation email timestamp
- `reason` - Cancellation reason
- `retailer` - Store name

---

## Adding New Retailers

You can easily add support for new retailers by editing the `RETAILER_RULES` dictionary in `scraperModule.py`:

```python
RETAILER_RULES = {
    "your_retailer": {
        # Email identification patterns
        "ids": ["retailername", "retailer.com", "orders@retailer"],
        
        # Order number extraction patterns (regex)
        "order_patterns": [
            r"order\s*#?\s*(\d{8,15})",
            r"order\s*number[:\s]*([A-Z0-9-]{8,20})"
        ],
        
        # Cancellation detection keywords
        "cancel_indicators": ["canceled", "cancelled", "order cancelled"],
        
        # Shipping status keywords
        "shipped_indicators": ["shipped", "on the way", "dispatched"],
        
        # Delivery confirmation keywords  
        "delivered_indicators": ["delivered", "arrived", "completed"],
        
        # Tracking number patterns (optional)
        "tracking_patterns": [
            r"\b(1Z[0-9A-Z]{16})\b",  # UPS
            r"\b(\d{12,22})\b"        # FedEx/Generic
        ]
    }
}
```

### Pattern Examples

**Order Number Patterns:**
- `r"order\s*#?\s*(\d{8,15})"` - Matches "Order #12345678"
- `r"\b([A-Z]{2}\d{8})\b"` - Matches "AB12345678"
- `r"order\s*number[:\s]*([A-Z0-9-]{10,20})"` - Matches "Order Number: ABC-123-DEF"

**Tracking Patterns:**
- `r"\b(1Z[0-9A-Z]{16})\b"` - UPS tracking (starts with 1Z)
- `r"\b(\d{12})\b"` - FedEx tracking (12 digits)
- `r"\b(\d{20,22})\b"` - USPS tracking (20-22 digits)

---

## GUI Features

### Configuration Tab
- **Add/Remove Accounts**: Manage multiple email accounts
- **Settings**: Configure scan period and output location
- **Save Configuration**: Persist settings for future use

### Run Scraper Tab
- **Real-time Logging**: See exactly what the scraper is doing
- **Progress Bar**: Visual progress indicator
- **Start/Stop Controls**: Full control over the scraping process

### View Results Tab
- **Results Overview**: Summary statistics and file locations
- **Sortable Table**: Click column headers to sort (3-state: unsorted → ascending → descending)
- **Context Menu**: Right-click for additional options
  - Hide orders from view
  - Copy tracking numbers to clipboard
- **Selection Counter**: Shows how many orders are selected
- **Refresh Results**: Update the display with latest data

### Advanced Features

**Smart Status Progression:**
Orders automatically advance through logical states:
- `ORDERED` → `SHIPPED` → `DELIVERED`
- `CANCELLED` (terminal state)
- Delivered orders never regress to earlier states

**Duplicate Prevention:**
- Cached emails prevent reprocessing
- Existing orders are updated, not duplicated

**Multi-Account Support:**
- Each account maintains separate cache files
- Results are merged into unified CSV reports

---

## Troubleshooting

### Common Issues

**Authentication Errors:**
- Ensure IMAP is enabled in your email settings
- Use App Passwords for accounts with 2FA
- Check IMAP server address

**No Orders Found:**
- Verify retailer emails are in your inbox (not spam/promotions)
- Increase the `DAYS_BACK` setting
- Check that order emails contain recognizable patterns

**Missing Tracking Numbers:**
- Some retailers send tracking info in separate emails
- Tracking patterns may need adjustment for specific carriers

**GUI Not Responding:**
- Large email accounts may take time to process
- Check the log output for progress updates
- Use the Stop button if needed

---

## Security Notes

- **App Passwords**: Use app-specific passwords instead of your main email password
- **Local Storage**: All data is stored locally on your computer
- **No Cloud Services**: No data is transmitted to external services
- **Cache Files**: Contain email metadata but not full email content

---

For support or questions, please create an issue in the repository with detailed information about your setup and any error messages.

---

## License

MIT — Feel free to use, modify, and share this tool.