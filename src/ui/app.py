import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import shutil
import tempfile
import json
import pandas as pd
from src.backend.invoices import InvoiceProcessor
from src.backend.statements import StatementProcessor
from src.config.config import load_settings, save_settings
import src.backend.constants as c




class DocumentProcessorApp(ttk.Window):
    def __init__(self):
        super().__init__(themename="superhero")  # You can change the theme here
        self.config_path = "src/config/setting.json"
        self.config = load_settings(self.config_path)
        print(self.config['excel'])

        self.title("Document Processor")
        self.geometry("400x400")
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
        self.process_button.pack_forget() 

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

            self.config['input'] = os.path.dirname(filenames[0])

            self.upload_button.pack_forget()
            self.process_button.pack(pady=20)
        else:
            self.status_label.config(text="No file selected.")
        
    def process_documents(self):
        if not self.filenames:
            messagebox.showwarning("Warning", "No files uploaded.")
            return
        
        self.temp_dir = tempfile.mkdtemp()

        if self.doc_type.get() == "statement":
            self.processor = StatementProcessor(self.config)
        else:
            self.processor = InvoiceProcessor(self.config)

        self.processor.read_excel_to_dict()
        found_pdf = False
        for filename in self.filenames:
            if filename.endswith(".pdf"):
                file_path = filename
                self.processor.pdf(file_path, self.temp_dir)
                found_pdf = True
        
        if found_pdf:
            self.processor.find_client()
            
            if self.processor.unmatched:
                self.show_unmatched_entries()
            else:
                self.show_results()
            self.process_button.pack_forget()
            self.upload_button.pack(pady=20)
        else:
            self.status_label.self.config(text="No PDFs found.")
            self.process_button.pack_forget()
            self.upload_button.pack(pady=20)
    
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

        # Define column headings with bold style
        headings = ["Customer Number", "Email", "Date", "Total", "Send"]
        header_style = {"foreground": "white", "font": ("TkDefaultFont", 10, "bold")}
        cell_style = {"foreground": "white"}

        for col_idx, heading in enumerate(headings):
            ttk.Label(scrollable_frame, text=heading, **header_style).grid(row=0, column=col_idx, padx=5, pady=5)

        self.checkboxes = []
        for idx, item in enumerate(self.processor.final_list, start=1):
            ttk.Label(scrollable_frame, text=item['customer_num'], **cell_style).grid(row=idx, column=0, padx=5, pady=5)
            ttk.Label(scrollable_frame, text=item['email'], **cell_style).grid(row=idx, column=1, padx=5, pady=5)
            ttk.Label(scrollable_frame, text=item.get('statement_date', item.get('invoice_num', '')), **cell_style).grid(row=idx, column=2, padx=5, pady=5)
            ttk.Label(scrollable_frame, text=item['total'], **cell_style).grid(row=idx, column=3, padx=5, pady=5)

            var = tk.BooleanVar(value=item['send'])
            cb = ttk.Checkbutton(scrollable_frame, variable=var, bootstyle="round-toggle")
            cb.grid(row=idx, column=4, padx=5, pady=5)
            self.checkboxes.append((item, var))

        action_frame = ttk.Frame(result_window)
        action_frame.pack(fill=tk.X, pady=10)

        continue_button = ttk.Button(action_frame, text="Continue", command=lambda: self.proceed(result_window), bootstyle="success")
        continue_button.pack(side=tk.LEFT, padx=20)

        exit_button = ttk.Button(action_frame, text="Exit", command=lambda: self.exit(result_window), bootstyle="danger")
        exit_button.pack(side=tk.RIGHT, padx=20)

    def show_unmatched_entries(self):
        unmatched_window = ttk.Toplevel(self)
        unmatched_window.title("Update Unmatched Entries")
        unmatched_window.geometry("500x600")

        frame = ttk.Frame(unmatched_window)
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

        # Define column headings with bold style
        headings = ["Customer Number", "Date", "Total", "Email", "Send"]
        header_style = {"foreground": "white", "font": ("TkDefaultFont", 10, "bold")}
        cell_style = {"foreground": "white"}

        for col_idx, heading in enumerate(headings):
            ttk.Label(scrollable_frame, text=heading, **header_style).grid(row=0, column=col_idx, padx=5, pady=5)

        self.unmatched_vars = []
        for idx, item in enumerate(self.processor.unmatched, start=1):
            email_var = tk.StringVar(value=item[c.EMAIL_KEY])
            send_var = tk.BooleanVar(value=item[c.SEND_KEY])
            ttk.Label(scrollable_frame, text=item['customer_num'], **cell_style).grid(row=idx, column=0, padx=5, pady=5)
            ttk.Label(scrollable_frame, text=item.get('statement_date', item.get('invoice_num', '')), **cell_style).grid(row=idx, column=1, padx=5, pady=5)
            ttk.Label(scrollable_frame, text=item['total'], **cell_style).grid(row=idx, column=2, padx=5, pady=5)
            ttk.Entry(scrollable_frame, textvariable=email_var).grid(row=idx, column=3, sticky='ew', padx=5, pady=5)
            ttk.Checkbutton(scrollable_frame, variable=send_var, bootstyle="round-toggle").grid(row=idx, column=4, sticky='w', padx=5, pady=5)
            self.unmatched_vars.append((item, email_var, send_var))

        action_frame = ttk.Frame(unmatched_window)
        action_frame.pack(fill=tk.X, pady=10)

        update_button = ttk.Button(action_frame, text="Update and Save", command=lambda: self.update_unmatched_entries(unmatched_window), bootstyle="success")
        update_button.pack(side=tk.LEFT, padx=20)

        exit_button = ttk.Button(action_frame, text="Exit", command=unmatched_window.destroy, bootstyle="danger")
        exit_button.pack(side=tk.RIGHT, padx=20)

    def update_unmatched_entries(self, window):
        for item, email_var, send_var in self.unmatched_vars:
            new_email = email_var.get()
            if new_email != item[c.EMAIL_KEY]:  # Only update if the email field was changed
                item[c.EMAIL_KEY] = new_email
                self.update_email_file(item)
                self.processor.final_list.append(item)                

        window.destroy()
        self.status_label.config(text="Unmatched entries updated.")
        self.show_results()

    def update_email_file(self, item):
        """
        Updates the emailfile.xlsx with the new customer code-email address pair if the email field was updated.
        """
        try:
            df = pd.read_excel(self.config['excel'])
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel file: {self.config['excel']}")
            return

        customer_code = item['customer_num']
        email = item[c.EMAIL_KEY]

        # Check if the email field was updated
        if not any((df[self.config['excel_customer']] == customer_code) & (df[self.config['excel_email']] == email)):
            if customer_code not in df[self.config['excel_customer']].values:
                new_row = pd.DataFrame({self.config['excel_customer']: [customer_code], self.config['excel_email']: [email]})
                df = pd.concat([df, new_row], ignore_index=True)
            else:
                df.loc[df[self.config['excel_customer']] == customer_code, self.config['excel_email']] = email

            try:
                df.to_excel(self.config['excel'], index=False)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to update Excel file: {e}")

    def clear_input_file(self, input_dir):
            root = tk.Tk()
            root.withdraw()  # Hide the main Tkinter window

            clear = messagebox.askyesno("Confirm Deletion", "Would you like to permanently delete the input document?")
            root.destroy()  # Destroy the Tkinter root window

            if clear:
                self.processor.clear_directory(input_dir)

    def proceed(self, window):
        for item, var in self.checkboxes:
            item['send'] = var.get()

        if self.temp_dir:
            if self.config['send_email']:
                self.processor.check_and_send_documents(self.processor.final_list, self.temp_dir)
                messagebox.showinfo("Emails Sent", f'All emails flag to send have been sent')
            else:
                messagebox.showinfo("Emails Not Sent", 'send_email was set to False, emails were not sent')

            if self.config['backup_pdf']:
                self.processor.copy_and_clear_directory(self.temp_dir, self.config['backup'], self.doc_type.get())
                self.clear_input_file(self.config['input'])
                messagebox.showinfo("Backup Complete", f'The input directory has been cleared and today\'s documents have been backed up into: {self.config["backup"]}')
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

        container = ttk.Frame(settings_window)
        container.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(container)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.settings_vars = {}
        row = 0
        for key, value in self.config.items():
            if key not in ["input_statements", "input_invoices", "database_dir", "database_filename"]:
                ttk.Label(scrollable_frame, text=key).grid(row=row, column=0, sticky='w', pady=5)
                var = ttk.StringVar(value=str(value))
                entry = ttk.Entry(scrollable_frame, textvariable=var)
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
                self.config[key] = True
            elif value.lower() == 'false':
                self.config[key] = False
            else:
                try:
                    self.config[key] = int(value)
                except ValueError:
                    try:
                        self.config[key] = float(value)
                    except ValueError:
                        self.config[key] = value
        save_settings(self.config, self.config_path)
        messagebox.showinfo("Settings Saved", "self.configuration has been saved successfully.")
