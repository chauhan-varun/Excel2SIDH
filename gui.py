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
        self.excel_path_var = tk.StringVar()
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

        # Excel File Row
        ttk.Label(config_frame, text="Excel File Path:", style="Card.TLabel").grid(row=0, column=0, sticky='w', pady=5)
        excel_entry = tk.Entry(config_frame, textvariable=self.excel_path_var, bg=self.bg_color, fg=self.text_color, insertbackground=self.text_color, bd=1, relief='flat', highlightbackground=self.border_color, highlightcolor=self.accent_color, highlightthickness=1)
        excel_entry.grid(row=0, column=1, columnspan=2, sticky='ew', padx=(10, 10), pady=5)
        
        browse_btn = tk.Button(config_frame, text="Browse", command=self.browse_excel, bg="#333333", fg=self.text_color, activebackground="#444444", activeforeground=self.text_color, relief='flat', bd=0, padx=10, pady=2)
        browse_btn.grid(row=0, column=3, sticky='e', pady=5)

        # Batch ID Row
        ttk.Label(config_frame, text="Batch ID:", style="Card.TLabel").grid(row=1, column=0, sticky='w', pady=5)
        self.batch_entry = tk.Entry(config_frame, textvariable=self.batch_id_var, bg=self.bg_color, fg=self.text_color, insertbackground=self.text_color, bd=1, relief='flat', highlightbackground=self.border_color, highlightcolor=self.accent_color, highlightthickness=1)
        self.batch_entry.grid(row=1, column=1, sticky='w', padx=(10, 0), pady=5, width=15)

        # Start From Row
        ttk.Label(config_frame, text="Start From Row Index:", style="Card.TLabel").grid(row=1, column=2, sticky='w', pady=5, padx=(20, 0))
        start_entry = tk.Entry(config_frame, textvariable=self.start_from_var, bg=self.bg_color, fg=self.text_color, insertbackground=self.text_color, bd=1, relief='flat', highlightbackground=self.border_color, highlightcolor=self.accent_color, highlightthickness=1)
        start_entry.grid(row=1, column=3, sticky='w', pady=5, width=10)

        # Student Filter Row
        ttk.Label(config_frame, text="Specific Candidate ID (Optional):", style="Card.TLabel").grid(row=2, column=0, sticky='w', pady=5)
        student_entry = tk.Entry(config_frame, textvariable=self.student_filter_var, bg=self.bg_color, fg=self.text_color, insertbackground=self.text_color, bd=1, relief='flat', highlightbackground=self.border_color, highlightcolor=self.accent_color, highlightthickness=1)
        student_entry.grid(row=2, column=1, columnspan=2, sticky='w', padx=(10, 0), pady=5, width=25)

        # Dry Run Checkbox
        dry_run_check = ttk.Checkbutton(config_frame, text="Dry Run Mode (Test fill only, don't submit)", variable=self.dry_run_var)
        dry_run_check.grid(row=3, column=0, columnspan=3, sticky='w', pady=(10, 5))

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
        self.log_widget = ScrolledText(log_frame, bg="#0d0d0d", fg="#00ff00", insertbackground="#ffffff", relief='flat', bd=1, font=('Consolas', 9.5), wrap='word')
        self.log_widget.pack(fill='both', expand=True)
        self.log_widget.tag_configure("info", foreground="#00ff00")
        self.log_widget.tag_configure("warning", foreground="#ffaa00")
        self.log_widget.tag_configure("error", foreground="#ff3333")
        self.log_widget.tag_configure("success", foreground="#33ff33", font=('Consolas', 9.5, 'bold'))

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

    def browse_excel(self):
        filename = filedialog.askopenfilename(
            title="Select Excel Result Sheet",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.excel_path_var.set(filename)
            # Try to auto-extract batch ID from the name (e.g. 3391656)
            match = re.search(r'\d{6,8}', os.path.basename(filename))
            if match:
                self.batch_id_var.set(match.group(0))

    def toggle_automation(self):
        if self.is_running:
            self.stop_automation()
        else:
            self.start_automation()

    def start_automation(self):
        excel_path = self.excel_path_var.get().strip()
        batch_id = self.batch_id_var.get().strip()
        
        # Validation
        if not excel_path:
            messagebox.showerror("Error", "Please select an Excel sheet first.")
            return
        if not os.path.exists(excel_path):
            messagebox.showerror("Error", f"Excel file does not exist:\n{excel_path}")
            return
        if not batch_id:
            messagebox.showerror("Error", "Please enter a valid Batch ID.")
            return
            
        try:
            start_from = int(self.start_from_var.get().strip())
        except ValueError:
            messagebox.showerror("Error", "Start index must be a valid number.")
            return

        self.is_running = True
        self.action_btn.configure(text="Stop Automation", bg="#ff3333", activebackground="#cc2222")
        self.status_label.configure(text="Initializing browser...", foreground=self.accent_color)
        self.log_widget.configure(state='normal')
        self.log_widget.delete('1.0', 'end')
        self.log_widget.configure(state='disabled')

        # Launch the automation in a background thread to prevent GUI freezing
        self.thread = threading.Thread(target=self.run_process, args=(excel_path, batch_id, start_from))
        self.thread.daemon = True
        self.thread.start()

    def stop_automation(self):
        self.status_label.configure(text="Stopping...", foreground="#ff3333")
        self.is_running = False

    def run_process(self, excel_path, batch_id, start_from):
        global BATCH_ID, EXCEL_FILE, BATCH_PAGE_URL
        
        try:
            # Dynamically update the configuration global variables of automate.py
            import automate
            automate.BATCH_ID = batch_id
            automate.EXCEL_FILE = excel_path
            automate.BATCH_PAGE_URL = f"{BASE_URL}/admin-profile/assessor/master-assessor/view-batch-details-new/PENDING/{batch_id};batches=assessment"

            # Remove previous file handler if any exists
            if hasattr(self, 'file_handler') and self.file_handler:
                logging.getLogger().removeHandler(self.file_handler)
                self.file_handler.close()
                self.file_handler = None

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
                    self.finish_run("Student not found", False)
                    return

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
                        self.finish_run("Failed to install browser dependencies. Run 'playwright install chromium' manually.", False)
                        return
                else:
                    raise e

            self.root.after(0, lambda: self.status_label.configure(
                text=f"Browser active. Log in to portal manually.", foreground="#00aaff"
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

                if not is_logged_in:
                    logging.info("🔒 Portal not logged in. Waiting for assessor to log in manually...")
                    # Show an alert asking to log in
                    self.root.after(0, lambda: messagebox.showinfo(
                        "Login Required", 
                        "A browser window has opened.\n\n1. Please log in to the Skill India Portal manually.\n2. Once fully logged in and on the batch details page, click OK here to start the automation."
                    ))

                # Re-verify page context
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
                self.finish_run("Automation completed successfully!", True)

        except Exception as e:
            logging.error(f"Critical process failure: {e}")
            self.finish_run(f"Process crashed: {e}", False)

    def finish_run(self, message, is_success):
        # Remove and close file handler if it exists
        if hasattr(self, 'file_handler') and self.file_handler:
            logging.getLogger().removeHandler(self.file_handler)
            self.file_handler.close()
            self.file_handler = None

        def update():
            self.is_running = False
            self.action_btn.configure(text="Start Automation", bg=self.accent_color, activebackground=self.accent_hover)
            self.status_label.configure(text="Ready to start.", foreground=self.text_secondary)
            if is_success:
                messagebox.showinfo("Success", message)
            else:
                messagebox.showerror("Error", message)

        self.root.after(0, update)

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
