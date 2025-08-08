# Target Email Scraper

This tool connects to your email inbox, scans for Target emails, and extracts order information including both successful orders and cancellations. It provides detailed tracking of your Target shopping activity with automatic categorization.

## What it extracts:

For Orders:
- Order number  
- Tracking numbers (UPS, FedEx, USPS)  
- Ship-to address  
- Recipient email and sent date  
- Order status (ordered/shipped)

For Cancellations:
- Order number
- Cancellation reason (payment issues, purchase limits, etc.)
- Cancellation date and recipient email

## Features

- Two interfaces: Command-line script + modern GUI application
- Multiple email accounts: Process several Gmail accounts simultaneously
- Secure IMAP login using Gmail app passwords  
- Smart date filtering: Only scans emails from the past X days  
- Intelligent caching: Avoids re-processing the same emails
- Dual CSV output: Separate files for orders and cancellations
- Analytics: Cancellation rate tracking and order summaries
- Target-specific: Optimized for Target.com email patterns
- Clean extraction: Parses HTML emails with BeautifulSoup

## License

MIT — Feel free to use, modify, and share this tool.