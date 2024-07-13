import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import shutil
import tempfile
import json
from src.backend.invoices import InvoiceProcessor
from src.backend.statements import StatementProcessor
from src.config.config import load_settings, save_settings

config_path = "src/config/setting.json"
config = load_settings(config_path)

class DocumentProcessorApp(ttk.Window):
    def __init__(self):
        super().__init__(themename="superhero")  # You can change the theme here
        self.title("Document Processor")
        self.geometry("400x350")
        self.create_widgets()
        
    def create_widgets(self):
        self.doc_type = ttk.StringVar(value="statement")
        
        self.radio_statement = ttk.Radiobutton(self, text="Statements", variable=self.doc_type, value="statement", bootstyle="success")
        self.radio_statement.pack(pady=10)
        
        self.radio_invoice = ttk.Radiobutton(self, text="Invoices", variable=self.doc_type, value="invoice", bootstyle="info")
        self.radio_invoice.pack(pady=10)
        
        self.upload_button = ttk.Button(self, text="Upload PDF", command=self.upload_file, bootstyle="primary")
        self.upload_button.pack(pady=20)
        
        self.process_button = ttk.Button(self, text="Process Documents", command=self.process_documents, bootstyle="success")
        self.process_button.pack(pady=20)

        self.settings_button = ttk.Button(self, text="Settings", command=self.open_settings, bootstyle="warning")
        self.settings_button.pack(pady=20)
        
        self.status_label = ttk.Label(self, text="", wraplength=500)
        self.status_label.pack(pady=20)
        
        self.temp_dir = None
        self.processor = None
        self.filenames = []
        
    def upload_file(self):
        filetypes = (("PDF files", "*.pdf"), ("All files", "*.*"))
        filenames = filedialog.askopenfilenames(title="Select files", filetypes=filetypes)
        if filenames:
            self.filenames = filenames
            uploaded_files = ", ".join([os.path.basename(file) for file in self.filenames])
            self.status_label.config(text=f"Uploaded {len(self.filenames)} files: {uploaded_files}")
        else:
            self.status_label.config(text="No file selected.")
        
    def process_documents(self):
        if not self.filenames:
            messagebox.showwarning("Warning", "No files uploaded.")
            return
        
        self.temp_dir = tempfile.mkdtemp()
        
        
        if self.doc_type.get() == "statement":
            self.processor = StatementProcessor(config)
        else:
            self.processor = InvoiceProcessor(config)

        self.processor.read_excel_to_dict()
        found_pdf = False
        for filename in self.filenames:
            if filename.endswith(".pdf"):
                file_path = filename
                self.processor.pdf(file_path, self.temp_dir)
                found_pdf = True
        
        if found_pdf:
            self.processor.find_client()
            self.show_results()
        else:
            self.status_label.config(text="No PDFs found.")
    
    def show_results(self):
        result_window = ttk.Toplevel(self)
        result_window.title("Review Results")
        result_window.geometry("800x600")

        frame = ttk.Frame(result_window)
        frame.pack(expand=True, fill=tk.BOTH)

        canvas = tk.Canvas(frame)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        self.checkboxes = []
        for idx, item in enumerate(self.processor.final_list):
            var = tk.BooleanVar(value=item['send'])
            cb = ttk.Checkbutton(scrollable_frame, text=f"{item['customer_num']} - {item['email']} - {item.get('statement_date', item.get('invoice_num', ''))} - {item['total']}", variable=var, bootstyle="round-toggle")
            cb.grid(row=idx, sticky='w')
            self.checkboxes.append((item, var))

        action_frame = ttk.Frame(result_window)
        action_frame.pack(fill=tk.X, pady=10)

        continue_button = ttk.Button(action_frame, text="Continue", command=lambda: self.proceed(result_window), bootstyle="success")
        continue_button.pack(side=tk.LEFT, padx=20)

        exit_button = ttk.Button(action_frame, text="Exit", command=lambda: self.exit(result_window), bootstyle="danger")
        exit_button.pack(side=tk.RIGHT, padx=20)

    def proceed(self, window):
        for item, var in self.checkboxes:
            item['send'] = var.get()

        if self.temp_dir:
            if config['send_email']:
                self.processor.check_and_send_documents(self.processor.final_list, self.temp_dir)
            if config['backup_pdf']:
                self.processor.copy_and_clear_directory(self.temp_dir, config['backup'], self.doc_type.get())
                messagebox.showinfo("Backup Complete", f'The input directory has been cleared and today\'s documents have been backed up into: {config["backup"]}')
        window.destroy()
        self.status_label.config(text="Processing completed.")

    def exit(self, window):
        if self.temp_dir:
            shutil.rmtree(self.temp_dir)
        window.destroy()
        self.status_label.config(text="Temporary files deleted.")

    def open_settings(self):
        settings_window = ttk.Toplevel(self)
        settings_window.title("Settings")
        settings_window.geometry("600x600")

        settings_frame = ttk.Frame(settings_window)
        settings_frame.pack(expand=True, fill=tk.BOTH, padx=20, pady=20)

        self.settings_vars = {}
        row = 0
        for key, value in config.items():
            if key != "input":
                ttk.Label(settings_frame, text=key).grid(row=row, column=0, sticky='w', pady=5)
                var = ttk.StringVar(value=str(value))
                entry = ttk.Entry(settings_frame, textvariable=var)
                entry.grid(row=row, column=1, sticky='ew', pady=5)
                entry.config(width=60)
                self.settings_vars[key] = var
                row += 1

        save_button = ttk.Button(settings_window, text="Save", command=self.save_settings, bootstyle="primary")
        save_button.pack(pady=10)

    def save_settings(self):
        for key, var in self.settings_vars.items():
            value = var.get()
            # Convert to the appropriate type (bool, int, or float) if needed
            if value.lower() == 'true':
                config[key] = True
            elif value.lower() == 'false':
                config[key] = False
            else:
                try:
                    config[key] = int(value)
                except ValueError:
                    try:
                        config[key] = float(value)
                    except ValueError:
                        config[key] = value
        save_settings(config, config_path)
        messagebox.showinfo("Settings Saved", "Configuration has been saved successfully.")
