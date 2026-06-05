import os
import re
import sys
import time
import threading
import logging
import argparse
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

# Add project root to path so we can import from automate
sys.path.append(str(Path(__file__).parent.resolve()))
from automate import read_excel_data, sync_playwright, SIDHAutomation, BATCH_PAGE_URL, BASE_URL

# ─── Custom UI Logger ────────────────────────────────────────────────────────

class QueueHandler(logging.Handler):
    """Logging handler that redirects logs to a Tkinter text widget via a callback."""
    def __init__(self, callback):
        super().__init__()
        self.callback = callback

    def emit(self, record):
        log_entry = self.format(record)
        self.callback(log_entry)

# ─── Main GUI Class ──────────────────────────────────────────────────────────

class SIDHAutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Skill India Digital Hub - Marks Entry Automator")
        self.root.geometry("850x650")
        self.root.minsize(700, 500)
        
        # Color Palette (Dark Theme with Skill India Amber/Orange accent)
        self.bg_color = "#121212"
        self.card_color = "#1e1e1e"
        self.accent_color = "#f07f22"
        self.accent_hover = "#d96f1c"
        self.text_color = "#ffffff"
        self.text_secondary = "#aaaaaa"
        self.border_color = "#333333"

        # Apply dark theme styling
        self.setup_styles()
        
        # Main Layout
        self.root.configure(bg=self.bg_color)
        
        # Variables
        self.queue = []
        self.batch_id_var = tk.StringVar()
        self.dry_run_var = tk.BooleanVar(value=True)
        self.start_from_var = tk.StringVar(value="1")
        self.student_filter_var = tk.StringVar()
        
        self.is_running = False
        self.thread = None
        self.playwright_context = None

        self.create_widgets()
        
        # Setup logging redirection
        self.setup_logging()

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('default')
        
        # Configure frames and layouts
        style.configure('TFrame', background=self.bg_color)
        style.configure('Card.TFrame', background=self.card_color, relief='flat', borderwidth=1)
        
        # Labels
        style.configure('TLabel', background=self.bg_color, foreground=self.text_color, font=('Segoe UI', 10))
        style.configure('Card.TLabel', background=self.card_color, foreground=self.text_color, font=('Segoe UI', 10))
        style.configure('Title.TLabel', background=self.card_color, foreground=self.accent_color, font=('Segoe UI', 16, 'bold'))
        style.configure('Sub.TLabel', background=self.card_color, foreground=self.text_secondary, font=('Segoe UI', 9))
        
        # Entries
        style.configure('TEntry', fieldbackground=self.bg_color, foreground=self.text_color, bordercolor=self.border_color, lightcolor=self.border_color, darkcolor=self.border_color)
        
        # Buttons
        style.configure('TButton', font=('Segoe UI', 10, 'bold'), background=self.accent_color, foreground='#ffffff', borderwidth=0)
        style.map('TButton',
                  background=[('active', self.accent_hover), ('disabled', '#555555')],
                  foreground=[('disabled', '#888888')])
                  
        style.configure('Browse.TButton', font=('Segoe UI', 9), background='#333333', foreground=self.text_color)
        style.map('Browse.TButton', background=[('active', '#444444')])

        # Checkbutton
        style.configure('TCheckbutton', background=self.card_color, foreground=self.text_color, font=('Segoe UI', 10))
        style.map('TCheckbutton', background=[('active', self.card_color)], foreground=[('active', self.text_color)])

        # Treeview Styling for Dark Theme
        style.configure('Treeview',
                        background=self.card_color,
                        foreground=self.text_color,
                        fieldbackground=self.card_color,
                        rowheight=25,
                        font=('Segoe UI', 9),
                        bordercolor=self.border_color,
                        borderwidth=1)
        style.map('Treeview',
                  background=[('selected', self.accent_color)],
                  foreground=[('selected', '#ffffff')])
        
        style.configure('Treeview.Heading',
                        background='#222222',
                        foreground=self.text_color,
                        font=('Segoe UI', 9, 'bold'))
        style.map('Treeview.Heading',
                  background=[('active', '#333333')])

    def create_widgets(self):
        # Header / Title Bar
        header_frame = ttk.Frame(self.root, style='Card.TFrame')
        header_frame.pack(fill='x', padx=15, pady=(15, 10))
        header_frame.configure(padding=15)
        
        title_label = ttk.Label(header_frame, text="Skill India Digital Hub Automator", style="Title.TLabel")
        title_label.pack(anchor='w')
        sub_label = ttk.Label(header_frame, text="Automate assessor marks entry from Excel sheets securely", style="Sub.TLabel")
        sub_label.pack(anchor='w', pady=(2, 0))

        # Configuration Card
        config_frame = ttk.Frame(self.root, style='Card.TFrame')
        config_frame.pack(fill='x', padx=15, pady=5)
        config_frame.configure(padding=15)

        # Configure columns of config_frame (Queue on left, settings on right)
        config_frame.columnconfigure(0, weight=3)
        config_frame.columnconfigure(1, weight=2)

        # --- Left Side: File Queue ---
        queue_frame = ttk.Frame(config_frame, style='Card.TFrame')
        queue_frame.grid(row=0, column=0, sticky='nsew', padx=(0, 15))

        ttk.Label(queue_frame, text="Excel File Queue:", style="Card.TLabel", font=('Segoe UI', 10, 'bold')).pack(anchor='w', pady=(0, 5))

        # Treeview and scrollbar container
        tree_container = ttk.Frame(queue_frame, style='Card.TFrame')
        tree_container.pack(fill='both', expand=True)

        self.queue_tree = ttk.Treeview(tree_container, columns=("filename", "batch_id", "status"), show="headings", height=5, selectmode="browse")
        self.queue_tree.heading("filename", text="Excel File")
        self.queue_tree.heading("batch_id", text="Batch ID")
        self.queue_tree.heading("status", text="Status")

        self.queue_tree.column("filename", width=250, anchor='w')
        self.queue_tree.column("batch_id", width=90, anchor='center')
        self.queue_tree.column("status", width=90, anchor='center')

        tree_scroll = ttk.Scrollbar(tree_container, orient="vertical", command=self.queue_tree.yview)
        self.queue_tree.configure(yscrollcommand=tree_scroll.set)

        self.queue_tree.pack(side='left', fill='both', expand=True)
        tree_scroll.pack(side='right', fill='y')

        # Bind row selection
        self.queue_tree.bind("<<TreeviewSelect>>", self.on_queue_select)

        # Buttons for Queue Management
        btn_frame = ttk.Frame(queue_frame, style='Card.TFrame')
        btn_frame.pack(fill='x', pady=(10, 0))

        self.add_btn = tk.Button(btn_frame, text="Add File(s)", command=self.browse_excels, bg="#333333", fg=self.text_color, activebackground="#444444", activeforeground=self.text_color, relief='flat', bd=0, padx=12, pady=4)
        self.add_btn.pack(side='left', padx=(0, 5))

        self.remove_btn = tk.Button(btn_frame, text="Remove Selected", command=self.remove_selected, bg="#333333", fg=self.text_color, activebackground="#444444", activeforeground=self.text_color, relief='flat', bd=0, padx=12, pady=4)
        self.remove_btn.pack(side='left', padx=5)

        self.clear_btn = tk.Button(btn_frame, text="Clear Queue", command=self.clear_queue, bg="#333333", fg=self.text_color, activebackground="#444444", activeforeground=self.text_color, relief='flat', bd=0, padx=12, pady=4)
        self.clear_btn.pack(side='left', padx=5)

        # --- Right Side: Selected File Settings ---
        settings_frame = ttk.Frame(config_frame, style='Card.TFrame')
        settings_frame.grid(row=0, column=1, sticky='nsew')

        ttk.Label(settings_frame, text="Queue Item Settings:", style="Card.TLabel", font=('Segoe UI', 10, 'bold')).grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 10))

        # Batch ID Row
        ttk.Label(settings_frame, text="Batch ID:", style="Card.TLabel").grid(row=1, column=0, sticky='w', pady=5)
        self.batch_entry = tk.Entry(settings_frame, textvariable=self.batch_id_var, width=15, bg=self.bg_color, fg=self.text_color, insertbackground=self.text_color, bd=1, relief='flat', highlightbackground=self.border_color, highlightcolor=self.accent_color, highlightthickness=1)
        self.batch_entry.grid(row=1, column=1, sticky='w', padx=(10, 0), pady=5)
        self.batch_id_var.trace_add("write", self.on_batch_id_changed)

        # Start From Row
        ttk.Label(settings_frame, text="Start Row Index:", style="Card.TLabel").grid(row=2, column=0, sticky='w', pady=5)
        start_entry = tk.Entry(settings_frame, textvariable=self.start_from_var, width=10, bg=self.bg_color, fg=self.text_color, insertbackground=self.text_color, bd=1, relief='flat', highlightbackground=self.border_color, highlightcolor=self.accent_color, highlightthickness=1)
        start_entry.grid(row=2, column=1, sticky='w', padx=(10, 0), pady=5)

        # Student Filter Row
        ttk.Label(settings_frame, text="Candidate ID (Opt):", style="Card.TLabel").grid(row=3, column=0, sticky='w', pady=5)
        student_entry = tk.Entry(settings_frame, textvariable=self.student_filter_var, width=20, bg=self.bg_color, fg=self.text_color, insertbackground=self.text_color, bd=1, relief='flat', highlightbackground=self.border_color, highlightcolor=self.accent_color, highlightthickness=1)
        student_entry.grid(row=3, column=1, sticky='w', padx=(10, 0), pady=5)

        # Dry Run Checkbox
        dry_run_check = ttk.Checkbutton(settings_frame, text="Dry Run Mode (Test fill only)", variable=self.dry_run_var)
        dry_run_check.grid(row=4, column=0, columnspan=2, sticky='w', pady=(10, 5))

        # Control Row
        control_frame = ttk.Frame(self.root, style='TFrame')
        control_frame.pack(fill='x', padx=15, pady=(10, 5))
        
        self.action_btn = tk.Button(control_frame, text="Start Automation", command=self.toggle_automation, bg=self.accent_color, fg="#ffffff", activebackground=self.accent_hover, activeforeground="#ffffff", font=('Segoe UI', 11, 'bold'), relief='flat', bd=0, padx=20, pady=6)
        self.action_btn.pack(side='left')

        self.status_label = ttk.Label(control_frame, text="Ready to start.", foreground=self.text_secondary, font=('Segoe UI', 10, 'italic'))
        self.status_label.pack(side='left', padx=15)

        # Log Terminal Card
        log_frame = ttk.Frame(self.root, style='Card.TFrame')
        log_frame.pack(fill='both', expand=True, padx=15, pady=(5, 15))
        log_frame.configure(padding=15)

        ttk.Label(log_frame, text="Automation Logs", style="Card.TLabel", font=('Segoe UI', 10, 'bold')).pack(anchor='w', pady=(0, 5))
        
        # ScrolledText for logs
        self.log_widget = ScrolledText(log_frame, bg="#0d0d0d", fg="#00ff00", insertbackground="#ffffff", relief='flat', bd=1, font=('Consolas', 10), wrap='word')
        self.log_widget.pack(fill='both', expand=True)
        self.log_widget.tag_configure("info", foreground="#00ff00")
        self.log_widget.tag_configure("warning", foreground="#ffaa00")
        self.log_widget.tag_configure("error", foreground="#ff3333")
        self.log_widget.tag_configure("success", foreground="#33ff33", font=('Consolas', 10, 'bold'))

        # Grid configuration for config frame
        config_frame.columnconfigure(1, weight=1)

    def setup_logging(self):
        """Redirect Python logger output directly into the ScrolledText widget."""
        formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S")
        
        # Custom emitter function that safely appends text in the GUI thread
        def append_log(msg):
            def write():
                self.log_widget.configure(state='normal')
                
                # Check message type to color-code it
                tag = "info"
                if "[WARNING]" in msg:
                    tag = "warning"
                elif "[ERROR]" in msg:
                    tag = "error"
                elif "✅" in msg or "Successfully" in msg or "verified" in msg:
                    tag = "success"
                
                self.log_widget.insert('end', msg + "\n", tag)
                self.log_widget.configure(state='disabled')
                self.log_widget.see('end')
            
            # Use root.after to run write() safely on the main thread
            self.root.after(0, write)

        handler = QueueHandler(append_log)
        handler.setFormatter(formatter)
        
        # Attach to root logger
        root_logger = logging.getLogger()
        root_logger.addHandler(handler)

    def browse_excels(self):
        filenames = filedialog.askopenfilenames(
            title="Select Excel Result Sheets",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filenames:
            for filename in filenames:
                # Check if already in queue
                if any(item["file_path"] == filename for item in self.queue):
                    continue
                
                # Auto-extract batch ID (e.g. 3391656)
                batch_id = ""
                match = re.search(r'\d{6,8}', os.path.basename(filename))
                if match:
                    batch_id = match.group(0)
                
                # Add to queue data source
                item = {
                    "file_path": filename,
                    "batch_id": batch_id,
                    "status": "Pending"
                }
                self.queue.append(item)
                
                # Insert to Treeview
                display_name = os.path.basename(filename)
                item_id = self.queue_tree.insert("", "end", values=(display_name, batch_id, "Pending"))
                # Store treeview item_id references
                item["tree_id"] = item_id
            
            # Select the last added item if nothing selected
            if not self.queue_tree.selection() and self.queue:
                self.queue_tree.selection_set(self.queue[-1]["tree_id"])

    def remove_selected(self):
        selected = self.queue_tree.selection()
        if not selected:
            return
        
        for item_id in selected:
            self.queue_tree.delete(item_id)
            self.queue = [item for item in self.queue if item.get("tree_id") != item_id]
            
        # Select another item or clear variables
        if self.queue:
            self.queue_tree.selection_set(self.queue[0]["tree_id"])
        else:
            self.batch_id_var.set("")

    def clear_queue(self):
        for item in self.queue:
            if "tree_id" in item:
                self.queue_tree.delete(item["tree_id"])
        self.queue.clear()
        self.batch_id_var.set("")

    def on_queue_select(self, event):
        selected = self.queue_tree.selection()
        if not selected:
            return
        
        item_id = selected[0]
        # Find item in self.queue
        for item in self.queue:
            if item.get("tree_id") == item_id:
                # Load its batch ID into the variable
                # Temporary disable trace to avoid infinite loop updating
                self._block_trace = True
                self.batch_id_var.set(item["batch_id"])
                self._block_trace = False
                break

    def on_batch_id_changed(self, *args):
        if getattr(self, "_block_trace", False):
            return
        
        selected = self.queue_tree.selection()
        if not selected:
            return
        
        item_id = selected[0]
        new_batch_id = self.batch_id_var.get().strip()
        
        # Update self.queue entry
        for item in self.queue:
            if item.get("tree_id") == item_id:
                item["batch_id"] = new_batch_id
                # Update Treeview text
                self.queue_tree.set(item_id, column="batch_id", value=new_batch_id)
                break

    def toggle_automation(self):
        if self.is_running:
            self.stop_automation()
        else:
            self.start_automation()

    def start_automation(self):
        if not self.queue:
            messagebox.showerror("Error", "Please add Excel sheets to the queue first.")
            return
            
        # Validation of Batch IDs in the queue
        for idx, item in enumerate(self.queue):
            if not item["batch_id"]:
                messagebox.showerror("Error", f"Item #{idx+1} ({os.path.basename(item['file_path'])}) has no Batch ID.\nPlease select it and set a Batch ID.")
                self.queue_tree.selection_set(item["tree_id"])
                return
            
        try:
            start_from = int(self.start_from_var.get().strip())
        except ValueError:
            messagebox.showerror("Error", "Start index must be a valid number.")
            return

        self.is_running = True
        
        # Disable inputs and buttons
        self.action_btn.configure(text="Stop Automation", bg="#ff3333", activebackground="#cc2222")
        self.add_btn.configure(state="disabled")
        self.remove_btn.configure(state="disabled")
        self.clear_btn.configure(state="disabled")
        self.batch_entry.configure(state="disabled")
        
        self.status_label.configure(text="Initializing queue...", foreground=self.accent_color)
        self.log_widget.configure(state='normal')
        self.log_widget.delete('1.0', 'end')
        self.log_widget.configure(state='disabled')

        # Launch the automation queue in a background thread
        self.thread = threading.Thread(target=self.run_queue, args=(start_from,))
        self.thread.daemon = True
        self.thread.start()

    def stop_automation(self):
        self.status_label.configure(text="Stopping...", foreground="#ff3333")
        self.is_running = False

    def update_item_status(self, tree_id, status):
        def update():
            self.queue_tree.set(tree_id, column="status", value=status)
        self.root.after(0, update)

    def run_queue(self, start_from):
        total_files = len(self.queue)
        success_count = 0
        failed_count = 0
        
        for index, item in enumerate(self.queue):
            if not self.is_running:
                break
                
            file_path = item["file_path"]
            batch_id = item["batch_id"]
            tree_id = item["tree_id"]
            
            # Select the item in the treeview so user sees what is running
            def select_item(tid=tree_id):
                self.queue_tree.selection_set(tid)
                self.queue_tree.see(tid)
            self.root.after(0, select_item)
            
            # Set status to Running
            self.update_item_status(tree_id, "Running")
            
            # Update overall progress status label
            filename = os.path.basename(file_path)
            self.root.after(0, lambda idx=index+1, name=filename: self.status_label.configure(
                text=f"File {idx} of {total_files}: {name}", foreground=self.accent_color
            ))
            
            # Log queue entry
            logging.info(f"\nQueue Progress: Processing {index+1}/{total_files} | File: {filename} (Batch ID: {batch_id})")
            
            # Execute the process for this file
            success = self.run_process(file_path, batch_id, start_from)
            
            if success:
                self.update_item_status(tree_id, "Success")
                success_count += 1
            else:
                self.update_item_status(tree_id, "Failed")
                failed_count += 1
                
            # For subsequent files, start_from resets to 1 (usually you only want start_from for the first file if resumed)
            start_from = 1
            
        # Re-enable inputs
        def enable_ui():
            self.is_running = False
            self.action_btn.configure(text="Start Automation", bg=self.accent_color, activebackground=self.accent_hover)
            self.status_label.configure(text="Ready to start.", foreground=self.text_secondary)
            self.add_btn.configure(state="normal")
            self.remove_btn.configure(state="normal")
            self.clear_btn.configure(state="normal")
            self.batch_entry.configure(state="normal")
            
            summary_msg = f"Queue finished.\n\nProcessed: {total_files}\nSuccess: {success_count}\nFailed: {failed_count}"
            messagebox.showinfo("Queue Finished", summary_msg)
            
        self.root.after(0, enable_ui)

    def close_file_handler(self):
        if hasattr(self, 'file_handler') and self.file_handler:
            logging.getLogger().removeHandler(self.file_handler)
            self.file_handler.close()
            self.file_handler = None

    def run_process(self, excel_path, batch_id, start_from):
        global BATCH_PAGE_URL
        
        try:
            # Dynamically update the configuration global variables of automate.py
            import automate
            automate.BATCH_ID = batch_id
            automate.EXCEL_FILE = excel_path
            automate.BATCH_PAGE_URL = f"{BASE_URL}/admin-profile/assessor/master-assessor/view-batch-details-new/PENDING/{batch_id};batches=assessment"

            self.close_file_handler()

            # Add FileHandler dynamically for this batch ID
            log_file = f"automation_{batch_id}.log"
            self.file_handler = logging.FileHandler(log_file, encoding="utf-8")
            self.file_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
            logging.getLogger().addHandler(self.file_handler)

            students = read_excel_data(excel_path)
            
            # Filter if specific student is requested
            student_filter = self.student_filter_var.get().strip()
            if student_filter:
                students = [s for s in students if s.enrollment_no == student_filter]
                if not students:
                    logging.error(f"Student {student_filter} not found in Excel sheet.")
                    self.close_file_handler()
                    return False

            # Ensure chromium is installed (crucial for PyInstaller packaging)
            try:
                with sync_playwright() as p:
                    p.chromium.launch(headless=True)
            except Exception as e:
                err_msg = str(e).lower()
                if "executable doesn't exist" in err_msg or "playwright install" in err_msg:
                    logging.info("🌐 Chromium browser not found. Starting first-time automatic download...")
                    self.root.after(0, lambda: self.status_label.configure(
                        text="Downloading browser (first-time setup)...", foreground=self.accent_color
                    ))
                    import subprocess
                    try:
                        # PyInstaller compiled executable forwards command-line flags
                        subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=True, capture_output=True)
                        logging.info("✅ Chromium browser downloaded and installed successfully!")
                    except Exception as download_err:
                        logging.error(f"Failed to automatically download Chromium: {download_err}")
                        self.close_file_handler()
                        return False
                else:
                    raise e

            self.root.after(0, lambda: self.status_label.configure(
                text="Browser active. Log in to portal manually.", foreground="#00aaff"
            ))

            with sync_playwright() as p:
                user_data_dir = Path(f"./.playwright_session_{batch_id}").resolve()
                logging.info(f"Launching Chromium persistent session at {user_data_dir}")
                
                context = p.chromium.launch_persistent_context(
                    user_data_dir=str(user_data_dir),
                    headless=False,
                    slow_mo=50,
                    viewport={"width": 1366, "height": 768},
                    user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                )
                self.playwright_context = context
                
                page = context.pages[0] if context.pages else context.new_page()
                
                # Navigate to the batch page
                logging.info(f"Navigating to: {automate.BATCH_PAGE_URL}")
                try:
                    page.goto(automate.BATCH_PAGE_URL, wait_until="domcontentloaded")
                except Exception as e:
                    logging.warning(f"Initial load redirected/slow: {e}")

                # Wait for manual login check
                is_logged_in = False
                # Try to check if we are already logged in
                for _ in range(3):
                    if not self.is_running:
                        break
                    for pge in context.pages:
                        try:
                            if pge.locator(f"text=Batch ID - {batch_id}").first.is_visible(timeout=1000) or \
                               pge.locator("text=Approved Applicants").first.is_visible(timeout=1000) or \
                               pge.locator("text=Approved Applicant").first.is_visible(timeout=1000):
                                is_logged_in = True
                                break
                        except Exception:
                            pass
                    if is_logged_in:
                        break
                    time.sleep(1)

                if not self.is_running:
                    context.close()
                    self.close_file_handler()
                    return False

                if not is_logged_in:
                    logging.info("🔒 Portal not logged in. Waiting for assessor to log in manually...")
                    # Show an alert asking to log in
                    login_ok = threading.Event()
                    def prompt_login():
                        messagebox.showinfo(
                            "Login Required", 
                            f"A browser window has opened for Batch {batch_id}.\n\n1. Please log in to the Skill India Portal manually.\n2. Once fully logged in and on the batch details page, click OK here to start the automation."
                        )
                        login_ok.set()
                    self.root.after(0, prompt_login)
                    # Wait for user to click OK
                    while not login_ok.is_set():
                        if not self.is_running:
                            context.close()
                            self.close_file_handler()
                            return False
                        time.sleep(0.5)

                # Re-verify page context
                if not self.is_running:
                    context.close()
                    self.close_file_handler()
                    return False

                from automate import get_active_page
                active_page = get_active_page(context)
                
                self.root.after(0, lambda: self.status_label.configure(
                    text="Running automation...", foreground=self.accent_color
                ))

                automation = SIDHAutomation(page=active_page, dry_run=self.dry_run_var.get())
                
                # Run the actual steps
                total = len(students)
                automation.navigate_to_batch()
                automation.click_approved_applicants_tab()

                for student in students:
                    if not self.is_running:
                        logging.info("🛑 Automation stopped by user.")
                        break

                    if student.serial_no < start_from:
                        continue

                    # Update status
                    self.root.after(0, lambda s=student: self.status_label.configure(
                        text=f"Entering marks: {s.name} ({s.serial_no}/{total})", foreground=self.accent_color
                    ))

                    result = automation.process_student(student, total)
                    automation.results.append(result)

                automation.print_summary()
                context.close()
                self.close_file_handler()
                return self.is_running

        except Exception as e:
            logging.error(f"Critical process failure: {e}")
            self.close_file_handler()
            return False

# ─── App Execution ──────────────────────────────────────────────────────────

def main():
    root = tk.Tk()
    
    # Enable rounded corners or native styling on modern systems
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
        
    app = SIDHAutomationApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
