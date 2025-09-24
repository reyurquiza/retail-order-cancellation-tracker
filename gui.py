import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import threading
import os
import json
import csv
from datetime import datetime
import sys
import queue
from scraperModule import scrape_target_emails
from config import EMAIL_ACCOUNTS, DAYS_BACK, CSV_PATH, CACHE_JSON, OUTPUT_DIR

def smart_trim(text, limit=90):
    if len(text) <= limit:
        return text

    trimmed = text[:limit].rsplit(" ", 1)[0]  # Cut to limit and backtrack to last full word
    return trimmed + "...\n"


class RealTimeLogger:
    """Custom logger that sends output to GUI in real-time"""

    def __init__(self, gui_log_queue):
        self.gui_log_queue = gui_log_queue
        self.original_stdout = sys.stdout
        self.original_stderr = sys.stderr

    def write(self, message):
        # Send to original stdout as well (for debugging)
        self.original_stdout.write(message)
        self.original_stdout.flush()

        # Send to GUI queue
        if message.strip():  # Only send non-empty messages
            self.gui_log_queue.put(message.strip())

    def flush(self):
        self.original_stdout.flush()


class TargetScraperGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Target Email Scraper")
        self.root.geometry("1100x700")
        self.root.minsize(1000, 600)

        # Configure style
        self.style = ttk.Style()
        self.style.theme_use('clam')

        # Create queue for real-time logging
        self.log_queue = queue.Queue()

        self.setup_ui()
        self.load_config()

        # Start checking for log messages
        self.check_log_queue()

    def check_log_queue(self):
        """Check for new log messages and display them"""
        try:
            while True:
                try:
                    message = self.log_queue.get_nowait()
                    timestamp = datetime.now().strftime("%H:%M:%S")
                    log_message = f"[{timestamp}] {message}\n"
                    self._append_log(smart_trim(log_message, limit=140))
                except queue.Empty:
                    break
        except Exception:
            pass  # Ignore errors in logging

        # Schedule next check
        self.root.after(100, self.check_log_queue)

    def setup_ui(self):
        # Create main notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Tab 1: Configuration
        self.config_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.config_frame, text="Configuration")
        self.setup_config_tab()

        # Tab 2: Run Scraper
        self.run_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.run_frame, text="Run Scraper")
        self.setup_run_tab()

        # Tab 3: View Results
        self.results_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.results_frame, text="View Results")
        self.setup_results_tab()

    def setup_config_tab(self):
        # Main container with scrollbar
        canvas = tk.Canvas(self.config_frame)
        scrollbar = ttk.Scrollbar(self.config_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Email Accounts Section
        accounts_label = ttk.Label(scrollable_frame, text="Email Accounts", font=('Arial', 14, 'bold'))
        accounts_label.pack(pady=(10, 5), anchor='w')

        self.accounts_frame = ttk.Frame(scrollable_frame)
        self.accounts_frame.pack(fill=tk.X, padx=10)

        self.account_entries = []

        # Add Account button
        add_button = ttk.Button(scrollable_frame, text="+ Add Account", command=self.add_account_entry)
        add_button.pack(pady=5)

        # Settings Section
        settings_label = ttk.Label(scrollable_frame, text="Settings", font=('Arial', 14, 'bold'))
        settings_label.pack(pady=(20, 5), anchor='w')

        settings_frame = ttk.Frame(scrollable_frame)
        settings_frame.pack(fill=tk.X, padx=10)

        # Days Back
        ttk.Label(settings_frame, text="Days to scan back:").grid(row=0, column=0, sticky='w', pady=5)
        self.days_back_var = tk.StringVar()
        ttk.Entry(settings_frame, textvariable=self.days_back_var, width=10).grid(row=0, column=1, sticky='w', padx=10)

        # Output Folder (single setting that controls CSV and cache locations)
        # TODO: FIX OUTPUT PATH IMPLEMENTATION
        # ttk.Label(settings_frame, text="Output folder:").grid(row=1, column=0, sticky='w', pady=5)
        self.output_dir_var = tk.StringVar()
        output_frame = ttk.Frame(settings_frame)
        output_frame.grid(row=1, column=1, sticky='ew', padx=10)
        #ttk.Entry(output_frame, textvariable=self.output_dir_var, width=50).pack(side=tk.LEFT, fill=tk.X, expand=True)
        #ttk.Button(output_frame, text="Browse", command=self.browse_output_folder).pack(side=tk.RIGHT, padx=(5, 0))

        settings_frame.columnconfigure(1, weight=1)

        # Save Configuration button
        save_button = ttk.Button(scrollable_frame, text="Save Configuration", command=self.save_config)
        save_button.pack(side=tk.BOTTOM, padx=20, pady=10)

        clear_cache_button = ttk.Button(scrollable_frame, text="Clear Cache", command=self.clear_cache)
        clear_cache_button.pack(side=tk.BOTTOM, padx=20, pady=10)

        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def setup_run_tab(self):
        # Control Panel
        control_frame = ttk.LabelFrame(self.run_frame, text="Control Panel", padding=10)
        control_frame.pack(fill=tk.X, padx=10, pady=10)

        button_frame = ttk.Frame(control_frame)
        button_frame.pack(fill=tk.X)

        self.start_button = ttk.Button(button_frame, text="Start Scraping", command=self.start_scraping,
                                       style='Accent.TButton')
        self.start_button.pack(side=tk.LEFT, padx=(0, 10))

        self.stop_button = ttk.Button(button_frame, text="Stop", command=self.stop_scraping, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT)

        # Progress bar
        self.progress_var = tk.StringVar(value="Ready to start...")
        ttk.Label(control_frame, textvariable=self.progress_var).pack(pady=(10, 0))

        self.progress_bar = ttk.Progressbar(control_frame, mode='indeterminate')
        self.progress_bar.pack(fill=tk.X, pady=5)

        # Log Output
        log_frame = ttk.LabelFrame(self.run_frame, text="Log Output", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=20)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Clear log button
        ttk.Button(log_frame, text="Clear Log", command=self.clear_log).pack(pady=(10, 0))

    def setup_results_tab(self):
        # Results overview
        overview_frame = ttk.LabelFrame(self.results_frame, text="Results Overview", padding=10)
        overview_frame.pack(fill=tk.X, padx=10, pady=10)

        self.results_text = tk.Text(overview_frame, height=6, state=tk.DISABLED)
        self.results_text.pack(fill=tk.X)

        # File buttons
        files_frame = ttk.LabelFrame(self.results_frame, text="Output Files", padding=10)
        files_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        ttk.Button(files_frame, text="Open Output Folder", command=self.open_orders_csv).pack(side=tk.LEFT,
                                                                                                 padx=(0, 10))
        ttk.Button(files_frame, text="Refresh Results", command=self.refresh_results).pack(side=tk.RIGHT)

        # Results table
        table_frame = ttk.LabelFrame(self.results_frame, text="Recent Orders", padding=10)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        # Create treeview for results (added Retailer and Tracking columns)
        columns = ("Order #", "Status", "Date", "Email", "Retailer", "Tracking")
        self.results_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15,
                                         selectmode='extended')

        # Add sorting functionality
        self.sort_reverse = {}  # Track sort direction for each column
        self.sort_states = {}  # Track sort state: 0=unsorted, 1=asc, 2=desc
        self.original_order = []  # Store original order for unsorted state
        self.hidden_items = set()  # Track hidden items

        for col in columns:
            self.results_tree.heading(col, text=col, command=lambda c=col: self.sort_treeview(c))
            # Give tracking and retailer column more width
            if col == "Tracking":
                self.results_tree.column(col, width=300)
            elif col == "Retailer":
                self.results_tree.column(col, width=120)
            elif col == "Email":
                self.results_tree.column(col, width=220)
            else:
                self.results_tree.column(col, width=120)
            self.sort_reverse[col] = False
            self.sort_states[col] = 0  # Start unsorted

        # Bind right-click context menu
        self.results_tree.bind("<Button-3>", self.show_context_menu)  # Right-click
        self.results_tree.bind("<<TreeviewSelect>>", self.update_selection_counter)

        # Create context menu
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Hide from view", command=self.hide_selected_items)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Copy tracking number(s)", command=self.copy_tracking_numbers)

        # Scrollbars for treeview
        tree_scrollbar_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.results_tree.yview)
        tree_scrollbar_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.results_tree.xview)
        self.results_tree.configure(yscrollcommand=tree_scrollbar_y.set, xscrollcommand=tree_scrollbar_x.set)

        self.results_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree_scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

        # Selection counter OUTSIDE and BELOW the table frame
        self.selection_counter = ttk.Label(self.results_frame, text="0 of 0 selected")
        self.selection_counter.pack(side=tk.LEFT, anchor='w', padx=10)

    def add_account_entry(self):
        account_frame = ttk.LabelFrame(self.accounts_frame, text=f"Account {len(self.account_entries) + 1}", padding=10)
        account_frame.pack(fill=tk.X, pady=5)

        # Email
        ttk.Label(account_frame, text="Email:").grid(row=0, column=0, sticky='w')
        email_var = tk.StringVar()
        ttk.Entry(account_frame, textvariable=email_var, width=30).grid(row=0, column=1, sticky='ew', padx=5)

        # Password
        ttk.Label(account_frame, text="Password:").grid(row=0, column=2, sticky='w', padx=(20, 0))
        password_var = tk.StringVar()
        ttk.Entry(account_frame, textvariable=password_var, show="*", width=20).grid(row=0, column=3, sticky='ew',
                                                                                     padx=5)

        # IMAP Server
        ttk.Label(account_frame, text="IMAP Server:").grid(row=1, column=0, sticky='w')
        imap_var = tk.StringVar(value="imap.gmail.com")
        ttk.Entry(account_frame, textvariable=imap_var, width=30).grid(row=1, column=1, sticky='ew', padx=5)

        # Remove button
        remove_button = ttk.Button(account_frame, text="Remove",
                                   command=lambda: self.remove_account_entry(account_frame))
        remove_button.grid(row=1, column=3, sticky='e', padx=5)

        account_frame.columnconfigure(1, weight=1)
        account_frame.columnconfigure(3, weight=1)

        self.account_entries.append({
            'frame': account_frame,
            'email': email_var,
            'password': password_var,
            'imap_server': imap_var
        })

    def remove_account_entry(self, frame):
        # Find and remove the entry
        for i, entry in enumerate(self.account_entries):
            if entry['frame'] == frame:
                frame.destroy()
                self.account_entries.pop(i)
                break

        # Renumber remaining frames
        for i, entry in enumerate(self.account_entries):
            entry['frame'].configure(text=f"Account {i + 1}")

    def browse_output_folder(self):
        folder = filedialog.askdirectory(title="Select output folder")
        if folder:
            self.output_dir_var.set(folder)

    def load_config(self):
        # Load from existing config
        try:
            self.days_back_var.set(str(DAYS_BACK))
            # Derive an output folder from CSV_PATH or CACHE_JSON if possible
            csv_dir = os.path.dirname(CSV_PATH) if CSV_PATH else ""
            cache_dir = os.path.dirname(CACHE_JSON) if CACHE_JSON else ""
            # Prefer CSV directory if it looks like a folder (non-empty)
            output_dir = csv_dir or cache_dir or ""
            if output_dir == "":
                output_dir = os.getcwd()
            self.output_dir_var.set(output_dir)

            # Load email accounts
            for account in EMAIL_ACCOUNTS:
                self.add_account_entry()
                entry = self.account_entries[-1]
                entry['email'].set(account.get('email', ''))
                entry['password'].set(account.get('password', ''))
                entry['imap_server'].set(account.get('imap_server', 'imap.gmail.com'))

        except Exception as e:
            self.log(f"Error loading config: {e}")

    def clear_cache(self):
        try:
            cache_path = OUTPUT_DIR + "/cache/"
            for filename in os.listdir(cache_path):
                file_path = os.path.join(cache_path, filename)
                os.remove(file_path)
                print(f"'{file_path}' has been deleted.")
        except OSError as e:
            print(f"Error deleting {file_path}: {e}")

    def save_config(self):
        try:
            # Prepare config data
            accounts = []
            for entry in self.account_entries:
                if entry['email'].get().strip():  # Only save non-empty emails
                    accounts.append({
                        'email': entry['email'].get().strip(),
                        'password': entry['password'].get(),
                        'imap_server': entry['imap_server'].get().strip()
                    })

            output_dir = self.output_dir_var.get().strip() or "output"
            # Make sure folder exists
            os.makedirs(output_dir, exist_ok=True)
            os.makedirs(os.path.join(output_dir, "cache"), exist_ok=True)

            # Create config content with proper path separators
            config_content = f'''# Multiple email accounts
import os
EMAIL_ACCOUNTS = {json.dumps(accounts, indent=4)}

DAYS_BACK = {self.days_back_var.get() or "7"}
OUTPUT_DIR = r"{output_dir}"
CSV_PATH = os.path.join(OUTPUT_DIR, "report.csv")
CACHE_JSON = os.path.join(OUTPUT_DIR, "cache", "emails.json")
    '''

            # Write to config.py
            with open('config.py', 'w') as f:
                f.write(config_content)

            # IMPORTANT: Reload the config module so changes take effect
            import importlib
            try:
                import config
                importlib.reload(config)
                self.log("Configuration reloaded successfully")
            except Exception as e:
                self.log(f"Warning: Could not reload config module: {e}")

            messagebox.showinfo("Success", "Configuration saved successfully!")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save configuration: {e}")

    def log(self, message):
        """Add message to log queue for real-time display"""
        self.log_queue.put(str(message))

    def _append_log(self, message):
        """Helper method to append log message on main thread"""
        self.log_text.insert(tk.END, message)
        self.log_text.see(tk.END)
        self.root.update_idletasks()  # Force immediate GUI update

    def clear_log(self):
        self.log_text.delete(1.0, tk.END)

    def start_scraping(self):
        if not self.account_entries:
            messagebox.showwarning("Warning", "Please add at least one email account.")
            return

        # Disable start button, enable stop button
        self.start_button.configure(state=tk.DISABLED)
        self.stop_button.configure(state=tk.NORMAL)
        self.progress_bar.start(10)

        # Clear log
        self.clear_log()

        # Start scraping in separate thread
        self.scraping_thread = threading.Thread(target=self.run_scraping, daemon=True)
        self.scraping_active = True
        self.scraping_thread.start()

    def stop_scraping(self):
        self.scraping_active = False
        self.progress_var.set("Stopping...")
        self.log("Stop requested by user...")

    def run_scraping(self):
        # Set up real-time logger
        real_time_logger = RealTimeLogger(self.log_queue)
        old_stdout = sys.stdout
        old_stderr = sys.stderr

        try:
            self.progress_var.set("Starting scraping process...")
            self.log("🚀 Starting Target email scraping...")

            # Get active accounts
            active_accounts = []
            for entry in self.account_entries:
                if entry['email'].get().strip():
                    active_accounts.append({
                        'email': entry['email'].get().strip(),
                        'password': entry['password'].get(),
                        'imap_server': entry['imap_server'].get().strip()
                    })

            if not active_accounts:
                self.log("❌ No active email accounts found!")
                return

            total_accounts = len(active_accounts)

            for i, account in enumerate(active_accounts, 1):
                if not self.scraping_active:
                    self.log("🛑 Scraping stopped by user")
                    break

                self.progress_var.set(f"Processing account {i}/{total_accounts}: {account['email']}")
                self.log(f"{'=' * 70}")
                self.log(f"Processing account {i}/{total_accounts}: {account['email']}")
                self.log(f"{'=' * 70}")

                try:
                    # Redirect stdout to real-time logger
                    sys.stdout = real_time_logger
                    sys.stderr = real_time_logger

                    # Run the scraper - output will now be shown in real-time
                    scrape_target_emails(
                        days_back=int(self.days_back_var.get() or "7"),
                        email_account=account['email'],
                        password=account['password'],
                        imap_server=account['imap_server']
                    )

                except Exception as e:
                    self.log(f"❌ Error processing {account['email']}: {e}")
                    continue
                finally:
                    # Restore original stdout/stderr
                    sys.stdout = old_stdout
                    sys.stderr = old_stderr

            if self.scraping_active:
                self.log(f"🎉 Finished processing all {total_accounts} accounts!")
                self.progress_var.set("Scraping completed successfully!")
            else:
                self.progress_var.set("Scraping stopped")

            # Refresh results
            self.root.after(1000, self.refresh_results)

        except Exception as e:
            self.log(f"❌ Critical error: {e}")
            self.progress_var.set("Error occurred during scraping")

        finally:
            # Restore original stdout/stderr
            sys.stdout = old_stdout
            sys.stderr = old_stderr

            # Re-enable buttons
            self.root.after(0, self.scraping_finished)

    def scraping_finished(self):
        self.start_button.configure(state=tk.NORMAL)
        self.stop_button.configure(state=tk.DISABLED)
        self.progress_bar.stop()
        self.scraping_active = False

    def show_context_menu(self, event):
        """Show right-click context menu"""
        try:
            # Select the item under cursor if not already selected
            item = self.results_tree.identify_row(event.y)
            if item and item not in self.results_tree.selection():
                self.results_tree.selection_set(item)

            selected_items = self.results_tree.selection()
            if not selected_items:
                return

            # Check if any selected items have tracking numbers
            has_tracking = self.check_tracking_numbers(selected_items)

            # Enable/disable tracking number option
            if has_tracking:
                self.context_menu.entryconfig("Copy tracking number(s)", state="normal")
            else:
                self.context_menu.entryconfig("Copy tracking number(s)", state="disabled")

            # Show context menu
            self.context_menu.post(event.x_root, event.y_root)

        except Exception as e:
            self.log(f"Error showing context menu: {e}")

    def check_tracking_numbers(self, selected_items):
        """Check if selected items have tracking numbers"""
        try:
            output_dir = self.output_dir_var.get().strip() or os.getcwd()
            orders_csv = os.path.join(output_dir, "report_orders.csv")

            if not os.path.exists(orders_csv):
                return False

            # Get order numbers from selected items
            selected_orders = []
            for item in selected_items:
                values = self.results_tree.item(item, 'values')
                if values:
                    selected_orders.append(values[0])  # Order number is first column

            # Check CSV for tracking numbers
            with open(orders_csv, 'r', newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    if row.get('order_number') in selected_orders:
                        tracking = row.get('tracking_numbers', '').strip()
                        if tracking:
                            return True

            return False

        except Exception as e:
            self.log(f"Error checking tracking numbers: {e}")
            return False

    def hide_selected_items(self):
        """Hide selected items from view"""
        try:
            selected_items = self.results_tree.selection()
            if not selected_items:
                return

            # Add to hidden items set
            for item in selected_items:
                values = self.results_tree.item(item, 'values')
                if values:
                    self.hidden_items.add(values[0])  # Add order number to hidden set

            # Remove from tree view
            for item in selected_items:
                self.results_tree.delete(item)

            self.update_selection_counter()
            self.log(f"Hidden {len(selected_items)} item(s) from view")

        except Exception as e:
            self.log(f"Error hiding items: {e}")

    def copy_tracking_numbers(self):
        """Copy tracking numbers from selected items to clipboard"""
        try:
            selected_items = self.results_tree.selection()
            if not selected_items:
                return

            output_dir = self.output_dir_var.get().strip() or os.getcwd()
            orders_csv = os.path.join(output_dir, "report_orders.csv")

            if not os.path.exists(orders_csv):
                messagebox.showwarning("Warning", "Orders CSV file not found.")
                return

            # Get order numbers from selected items
            selected_orders = []
            for item in selected_items:
                values = self.results_tree.item(item, 'values')
                if values:
                    selected_orders.append(values[0])  # Order number is first column

            # Collect tracking numbers
            tracking_numbers = []
            with open(orders_csv, 'r', newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    if row.get('order_number') in selected_orders:
                        tracking = row.get('tracking_numbers', '').strip()
                        if tracking:
                            # Split multiple tracking numbers if comma-separated
                            numbers = [num.strip() for num in tracking.split(',') if num.strip()]
                            tracking_numbers.extend(numbers)

            if tracking_numbers:
                # Copy to clipboard
                tracking_text = '\n'.join(tracking_numbers)
                self.root.clipboard_clear()
                self.root.clipboard_append(tracking_text)
                self.root.update()  # Ensure clipboard is updated

                self.log(f"Copied {len(tracking_numbers)} tracking number(s) to clipboard")
                messagebox.showinfo("Success", f"Copied {len(tracking_numbers)} tracking number(s) to clipboard")
            else:
                messagebox.showinfo("No Tracking", "No tracking numbers found for selected orders.")

        except Exception as e:
            self.log(f"Error copying tracking numbers: {e}")
            messagebox.showerror("Error", f"Failed to copy tracking numbers: {e}")

    def update_selection_counter(self, event=None):
        """Update the selection counter display"""
        try:
            selected_count = len(self.results_tree.selection())
            total_count = len(self.results_tree.get_children())

            self.selection_counter.config(text=f"{selected_count} of {total_count} selected")

        except Exception as e:
            self.log(f"Error updating selection counter: {e}")

    def sort_treeview(self, col):
        """Sort treeview contents when a column header is clicked - 3 states: unsorted, asc, desc"""
        try:
            # Cycle through sort states: 0 (unsorted) -> 1 (asc) -> 2 (desc) -> 0
            current_state = self.sort_states.get(col, 0)
            new_state = (current_state + 1) % 3
            self.sort_states[col] = new_state

            # Reset other columns to unsorted
            for other_col in self.sort_states.keys():
                if other_col != col:
                    self.sort_states[other_col] = 0
                    self.results_tree.heading(other_col, text=other_col)

            if new_state == 0:  # Unsorted - restore original order
                self.restore_original_order()
                self.results_tree.heading(col, text=col)
                self.log(f"Restored original order (unsorted)")

            else:  # Sorted states
                # Get all items and their values
                items = [(self.results_tree.set(child, col), child) for child in self.results_tree.get_children('')]

                reverse_sort = (new_state == 2)  # True for descending

                # Sort items based on column type
                if col == "Order #":
                    # Sort order numbers numerically
                    items.sort(key=lambda x: int(x[0]) if x[0].isdigit() else 0, reverse=reverse_sort)
                elif col == "Date":
                    # Sort dates chronologically
                    def parse_date(date_str):
                        try:
                            # Try to parse ISO format first
                            if 'T' in date_str:
                                return datetime.fromisoformat(date_str.replace('Z', '+00:00'))
                            # Try other common formats
                            for fmt in ['%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y']:
                                try:
                                    return datetime.strptime(date_str, fmt)
                                except ValueError:
                                    continue
                            return datetime.min
                        except:
                            return datetime.min

                    items.sort(key=lambda x: parse_date(x[0]), reverse=reverse_sort)
                else:
                    # Sort alphabetically for Status, Email, Retailer and Tracking
                    items.sort(key=lambda x: x[0].lower(), reverse=reverse_sort)

                # Rearrange items in sorted positions
                for index, (val, child) in enumerate(items):
                    self.results_tree.move(child, '', index)

                # Update column header to show sort direction
                direction = "▼" if reverse_sort else "▲"
                self.results_tree.heading(col, text=f"{col} {direction}")

                sort_type = "descending" if reverse_sort else "ascending"
                self.log(f"Sorted by {col} ({sort_type})")

        except Exception as e:
            self.log(f"Error sorting by {col}: {e}")

    def restore_original_order(self):
        """Restore the original order of items"""
        try:
            # If we have stored original order, restore it
            if hasattr(self, 'original_order') and self.original_order:
                current_items = {self.results_tree.set(child, "Order #"): child
                                 for child in self.results_tree.get_children('')}

                # Move items back to original order
                for index, order_num in enumerate(self.original_order):
                    if order_num in current_items:
                        child = current_items[order_num]
                        self.results_tree.move(child, '', index)
        except Exception as e:
            self.log(f"Error restoring original order: {e}")

    def refresh_results(self):
        try:
            # Clear existing results
            for item in self.results_tree.get_children():
                self.results_tree.delete(item)

            # Clear hidden items and reset sort states when refreshing
            self.hidden_items.clear()
            self.original_order = []

            # Reset all column headers and sort states
            for col in self.sort_states.keys():
                self.sort_states[col] = 0
                self.results_tree.heading(col, text=col)

            # Read and display results
            output_dir = self.output_dir_var.get().strip() or os.getcwd()
            orders_csv = os.path.join(output_dir, "report_orders.csv")
            cancellations_csv = os.path.join(output_dir, "report_cancellations.csv")

            total_orders = 0
            total_cancellations = 0

            # Read orders - ensure file is properly closed
            if os.path.exists(orders_csv):
                try:
                    with open(orders_csv, 'r', newline='', encoding='utf-8') as f:
                        reader = csv.DictReader(f)
                        orders_data = list(reader)  # Read all data at once

                    # Process data after file is closed
                    for row in orders_data:
                        order_num = row.get('order_number', '')
                        if order_num not in self.hidden_items:  # Only show non-hidden items
                            total_orders += 1
                            self.results_tree.insert('', 'end', values=(
                                order_num,
                                row.get('status', '').upper(),
                                row.get('sent_date', ''),
                                row.get('sent_to', ''),
                                row.get('retailer', ''),
                                row.get('tracking_numbers', '')
                            ))
                            # Store original order for unsorted state
                            self.original_order.append(order_num)
                except Exception as e:
                    self.log(f"Error reading orders CSV: {e}")

            # Read cancellations - ensure file is properly closed
            if os.path.exists(cancellations_csv):
                try:
                    with open(cancellations_csv, 'r', newline='', encoding='utf-8') as f:
                        reader = csv.DictReader(f)
                        cancellations_data = list(reader)  # Read all data at once

                    # Process data after file is closed
                    for row in cancellations_data:
                        order_num = row.get('order_number', '')
                        if order_num not in self.hidden_items:  # Only show non-hidden items
                            total_cancellations += 1
                            self.results_tree.insert('', 'end', values=(
                                order_num,
                                'CANCELLED',
                                row.get('sent_date', ''),
                                row.get('sent_to', ''),
                                row.get('retailer', ''),
                                ''  # cancellations don't have tracking in this table (or keep blank)
                            ))
                            # Store original order for unsorted state
                            self.original_order.append(order_num)
                except Exception as e:
                    self.log(f"Error reading cancellations CSV: {e}")

            # Update overview
            total_all = total_orders + total_cancellations
            cancellation_rate = (total_cancellations / total_all * 100) if total_all > 0 else 0

            overview_text = f"""Total Orders: {total_all}
Total Cancellations: {total_cancellations}
Cancellation Rate: {cancellation_rate:.1f}%

Orders CSV: {orders_csv}
Cancellations CSV: {cancellations_csv}

Last refreshed: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
"""

            self.results_text.configure(state=tk.NORMAL)
            self.results_text.delete(1.0, tk.END)
            self.results_text.insert(1.0, overview_text)
            self.results_text.configure(state=tk.DISABLED)

            # Update selection counter
            self.update_selection_counter()

            self.log(f"Results refreshed: {total_orders} orders, {total_cancellations} cancellations")

        except Exception as e:
            error_msg = f"Failed to refresh results: {e}"
            self.log(error_msg)
            messagebox.showerror("Error", error_msg)

    def open_orders_csv(self):
        output_dir = self.output_dir_var.get().strip() or os.getcwd()

        # Create directory if it doesn't exist
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir, exist_ok=True)
                self.log(f"Created output directory: {output_dir}")
            except Exception as e:
                messagebox.showerror("Error", f"Could not create directory: {e}")
                return

        # Open the directory
        try:
            if os.name == 'nt':  # Windows
                os.startfile(output_dir)
            elif os.name == 'posix':  # macOS and Linux
                import subprocess
                if sys.platform == 'darwin':  # macOS
                    subprocess.run(['open', output_dir])
                else:  # Linux
                    subprocess.run(['xdg-open', output_dir])
            self.log(f"Opened directory: {output_dir}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not open directory: {e}\n\nDirectory path: {output_dir}")

    def open_cancellations_csv(self):
        # Since both files are in the same directory, just call open_orders_csv
        self.open_orders_csv()


def main():
    root = tk.Tk()
    app = TargetScraperGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
