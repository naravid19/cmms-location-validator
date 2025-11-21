import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import logging
import json
import os
import sys

import queue

# Import the validator logic
import location_validator

# Configuration for the GUI
SETTINGS_FILE = "settings.json"
APP_NAME = "Location Validator"
THEME_COLOR = "dark-blue"  # customtkinter theme
GITHUB_URL = "https://github.com/naravid19/cmms-location-validator"
ICON_PATH = "checkmark.ico"
VERSION = "v1.0.0"
COPYRIGHT = "Copyright © 2025 Narawit"

class QueueHandler(logging.Handler):
    """This class sends log records to a queue"""
    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        self.log_queue.put(record)

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title(APP_NAME)
        self.geometry("700x600")
        ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
        ctk.set_default_color_theme(THEME_COLOR)
        
        # Queue for thread-safe logging
        self.log_queue = queue.Queue()
        
        if os.path.exists(ICON_PATH):
            self.iconbitmap(ICON_PATH)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # Start polling the log queue
        self.after(100, self.check_log_queue)

        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_columnconfigure(1, weight=1)

        # --- UI Elements ---
        
        # Title
        self.label_title = ctk.CTkLabel(self.main_frame, text="Location Validator", font=ctk.CTkFont(size=24, weight="bold"))
        self.label_title.grid(row=0, column=0, columnspan=3, padx=10, pady=(10, 20))

        # Sheet Name
        self.label_sheet = ctk.CTkLabel(self.main_frame, text="Sheet Name:")
        self.label_sheet.grid(row=1, column=0, padx=10, pady=10, sticky="e")
        self.entry_sheet = ctk.CTkEntry(self.main_frame, placeholder_text="e.g., LTK-H")
        self.entry_sheet.grid(row=1, column=1, columnspan=2, padx=10, pady=10, sticky="ew")

        # Input File
        self.label_input = ctk.CTkLabel(self.main_frame, text="Input File:")
        self.label_input.grid(row=2, column=0, padx=10, pady=10, sticky="e")
        self.entry_input = ctk.CTkEntry(self.main_frame, placeholder_text="Path to input .xlsm or .xlsx")
        self.entry_input.grid(row=2, column=1, padx=10, pady=10, sticky="ew")
        self.btn_input = ctk.CTkButton(self.main_frame, text="Browse", width=80, command=self.browse_input)
        self.btn_input.grid(row=2, column=2, padx=10, pady=10)

        # Database Code
        self.label_db = ctk.CTkLabel(self.main_frame, text="Database Code:")
        self.label_db.grid(row=3, column=0, padx=10, pady=10, sticky="e")
        self.entry_db = ctk.CTkEntry(self.main_frame, placeholder_text="Path to Database_Code.xlsx")
        self.entry_db.grid(row=3, column=1, padx=10, pady=10, sticky="ew")
        self.btn_db = ctk.CTkButton(self.main_frame, text="Browse", width=80, command=self.browse_db)
        self.btn_db.grid(row=3, column=2, padx=10, pady=10)

        # Run Button
        self.btn_run = ctk.CTkButton(self.main_frame, text="Run Validation", height=40, font=ctk.CTkFont(size=16, weight="bold"), command=self.start_validation)
        self.btn_run.grid(row=4, column=0, columnspan=3, padx=10, pady=20, sticky="ew")

        # Log Output
        self.textbox_log = ctk.CTkTextbox(self.main_frame, height=200)
        self.textbox_log.grid(row=5, column=0, columnspan=3, padx=10, pady=(0, 10), sticky="nsew")
        self.textbox_log.configure(state='disabled')
        self.main_frame.grid_rowconfigure(5, weight=1)
        
        # Footer Frame
        self.footer_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.footer_frame.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        self.footer_frame.grid_columnconfigure(0, weight=1)
        self.footer_frame.grid_columnconfigure(2, weight=1)

        # GitHub Link
        self.btn_github = ctk.CTkButton(self.footer_frame, text="GitHub Repo", width=100, fg_color="transparent", border_width=1, text_color=("gray10", "gray90"), command=self.open_github)
        self.btn_github.grid(row=0, column=0, sticky="w", padx=10)

        # Version & Copyright
        self.label_version = ctk.CTkLabel(self.footer_frame, text=f"{VERSION} | {COPYRIGHT}", font=ctk.CTkFont(size=10))
        self.label_version.grid(row=0, column=2, sticky="e", padx=10)

        # --- Setup Logging ---
        self.setup_logging()

        # --- Load Settings ---
        self.load_settings()

    def setup_logging(self):
        # Create queue handler
        queue_handler = QueueHandler(self.log_queue)
        queue_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        
        # Add to location_validator logger
        validator_logger = logging.getLogger('location_validator')
        
        # Clear existing handlers to prevent duplicates
        if validator_logger.hasHandlers():
            validator_logger.handlers.clear()
            
        validator_logger.addHandler(queue_handler)
        validator_logger.setLevel(logging.INFO)
        
    def check_log_queue(self):
        """Poll the queue for new log records and display them"""
        while not self.log_queue.empty():
            try:
                record = self.log_queue.get_nowait()
                msg = self.format_log_record(record)
                self.textbox_log.configure(state='normal')
                self.textbox_log.insert(tk.END, msg + '\n')
                self.textbox_log.configure(state='disabled')
                self.textbox_log.see(tk.END)
            except queue.Empty:
                break
        # Schedule next check
        self.after(100, self.check_log_queue)

    def format_log_record(self, record):
        # Manually format since we aren't using the handler's formatter directly in the loop
        # Or we can just use the formatter we created.
        # Let's keep it simple and match the previous format
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        return formatter.format(record)

    def open_github(self):
        webbrowser.open(GITHUB_URL)

    def load_settings(self):
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, 'r') as f:
                    settings = json.load(f)
                    
                if "sheet_name" in settings:
                    self.entry_sheet.insert(0, settings["sheet_name"])
                if "file_input" in settings:
                    self.entry_input.insert(0, settings["file_input"])
                if "database_code" in settings:
                    self.entry_db.insert(0, settings["database_code"])
                
                logging.info("Settings loaded.")
            except Exception as e:
                logging.error(f"Failed to load settings: {e}")

    def save_settings(self):
        settings = {
            "sheet_name": self.entry_sheet.get(),
            "file_input": self.entry_input.get(),
            "database_code": self.entry_db.get()
        }
        try:
            with open(SETTINGS_FILE, 'w') as f:
                json.dump(settings, f, indent=4)
            logging.info("Settings saved.")
        except Exception as e:
            logging.error(f"Failed to save settings: {e}")

    def browse_input(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xlsm")])
        if filename:
            self.entry_input.delete(0, tk.END)
            self.entry_input.insert(0, filename)

    def browse_db(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xlsm")])
        if filename:
            self.entry_db.delete(0, tk.END)
            self.entry_db.insert(0, filename)

    def start_validation(self):
        sheet_name = self.entry_sheet.get()
        file_input = self.entry_input.get()
        database_code = self.entry_db.get()

        if not sheet_name or not file_input or not database_code:
            messagebox.showwarning("Missing Input", "Please fill in all fields.")
            return

        if not os.path.exists(file_input):
            messagebox.showerror("Error", f"Input file not found: {file_input}")
            return
        
        if not os.path.exists(database_code):
            messagebox.showerror("Error", f"Database file not found: {database_code}")
            return

        # Save settings before running
        self.save_settings()

        # Disable button
        self.btn_run.configure(state="disabled", text="Processing...")
        self.textbox_log.configure(state='normal')
        self.textbox_log.delete('1.0', tk.END)
        self.textbox_log.configure(state='disabled')

        # Run in thread
        thread = threading.Thread(target=self.run_logic, args=(sheet_name, file_input, database_code))
        thread.start()

    def run_logic(self, sheet_name, file_input, database_code):
        try:
            success = location_validator.main(sheet_name, file_input, database_code)
            if success:
                self.after(0, lambda: messagebox.showinfo("Success", "Validation completed successfully!"))
            else:
                self.after(0, lambda: messagebox.showerror("Error", "Validation failed. Check logs for details."))
        except Exception as e:
            logging.error(f"Critical error in GUI thread: {e}")
            self.after(0, lambda: messagebox.showerror("Critical Error", str(e)))
        finally:
            self.after(0, self.reset_ui)

    def reset_ui(self):
        self.btn_run.configure(state="normal", text="Run Validation")

if __name__ == "__main__":
    app = App()
    app.mainloop()
