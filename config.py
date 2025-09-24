import os
# Multiple email accounts
EMAIL_ACCOUNTS = []


'''
Recommendations:
    1. First time running, use 365 days back to grab as many emails as possible, cancels will
    not be registered if an 'order' email is never detected.
    2. Keep Output_dir the default if possible.
'''
DAYS_BACK = 7
OUTPUT_DIR = "C:/Users/Administrator/Documents/retail-order-tracker/output"
CSV_PATH = os.path.join(OUTPUT_DIR, "report.csv")
CACHE_JSON = os.path.join(OUTPUT_DIR, "cache", "emails.json")
