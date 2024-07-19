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
        super().__init__(themename="litera")  # You can change the theme here
        self.config_path = "src/config/setting.json"
        self.config = load_settings(self.config_path)
        self.title("Document Processor")
        self.geometry("400x300")
        self.create_widgets()

    def create_widgets(self):
        self.doc_type = ttk.StringVar(value="statement")
        
        self.create_radiobuttons()
        self.create_buttons()
        self.create_status_label()

        self.temp_dir = None
        self.processor = None
        self.filenames = []

    def create_radiobuttons(self):
        # self.create_styles()
        radio_frame = ttk.Frame(self)
        radio_frame.pack(anchor="w", padx=20, pady=10)

        self.radio_statement = ttk.Radiobutton(radio_frame, text="Statements", variable=self.doc_type, value="statement", bootstyle="danger")
        self.radio_statement.pack(anchor="w", pady=10)

        self.last_invoice_run_text = tk.StringVar()
        self.update_invoice_text()
        self.radio_invoice = ttk.Radiobutton(radio_frame, text=self.last_invoice_run_text.get(), variable=self.doc_type, value="invoice", bootstyle="danger")
        self.radio_invoice.pack(anchor="w", pady=10)

    def create_buttons(self):
        button_frame = ttk.Frame(self)
        button_frame.pack(pady=20)

        self.settings_button = ttk.Button(button_frame, text="Settings", command=self.open_settings, bootstyle="danger")
        self.settings_button.pack(side=tk.LEFT, padx=10)

        self.upload_button = ttk.Button(button_frame, text="Upload PDF", command=self.upload_file, bootstyle="light")
        self.upload_button.pack(side=tk.LEFT, padx=10)

        self.process_button = ttk.Button(button_frame, text="Process Documents", command=self.process_documents, bootstyle="dark")
        self.process_button.pack(side=tk.LEFT, padx=10)
        self.process_button.pack_forget()

    def create_status_label(self):
        self.status_label = ttk.Label(self, text="", wraplength=500)
        self.status_label.pack(pady=20)

    def update_invoice_text(self):
        last_invoice_run = self.get_last_invoice_run()
        self.last_invoice_run_text.set(f"Invoices (Last run: {last_invoice_run})")

    def get_last_invoice_run(self):
        log_file = "logs/invoice_runs.log"
        try:
            with open(log_file, "r") as file:
                lines = file.readlines()
                if lines:
                    return lines[-1].strip().split("-")[-1]
        except FileNotFoundError:
            return "No previous runs logged."
        return "No previous runs logged."

    def upload_file(self):
        filetypes = (("PDF files", "*.pdf"), ("All files", "*.*"))
        filenames = filedialog.askopenfilenames(title="Select files", filetypes=filetypes)
        if filenames:
            self.filenames = filenames
            uploaded_files = ", ".join([os.path.basename(file) for file in self.filenames])
            self.status_label.config(text=f"Uploaded {len(self.filenames)} files: {uploaded_files}")

            self.config['input'] = os.path.dirname(filenames[0])

            self.upload_button.pack_forget()
            self.process_button.pack(padx=10)
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
        found_pdf = self.process_pdfs()

        if found_pdf:
            self.processor.find_client()
            if self.processor.unmatched:
                self.show_unmatched_entries()
            else:
                self.show_results()
            self.switch_to_upload()
        else:
            self.status_label.config(text="No PDFs found.")
            self.switch_to_upload()

    def process_pdfs(self):
        found_pdf = False
        for filename in self.filenames:
            if filename.endswith(".pdf"):
                file_path = filename
                success = self.processor.pdf(file_path, self.temp_dir)
                if success == "Error":
                    messagebox.showerror("Error", f"Invalid doc type uploaded or expect string not found, check the console for details")
                    return False
                found_pdf = True
        return found_pdf

    def switch_to_upload(self):
        self.process_button.pack_forget()
        self.upload_button.pack(pady=20)

    def show_results(self):
        result_window = self.create_window("Review Results", "700x600")
        scrollable_frame = self.create_scrollable_frame(result_window)

        self.create_result_headings(scrollable_frame, self.config["send_email"])
        self.populate_results(scrollable_frame, self.config["send_email"])

        self.create_action_buttons(result_window, self.proceed)

    def show_unmatched_entries(self):
        unmatched_window = self.create_window("Update Unmatched Entries", "500x600")
        scrollable_frame = self.create_scrollable_frame(unmatched_window)

        self.create_unmatched_headings(scrollable_frame)
        self.populate_unmatched(scrollable_frame)

        self.create_action_buttons(unmatched_window, self.update_unmatched_entries)

    def create_window(self, title, dimensions):
        window = ttk.Toplevel(self)
        window.title(title)
        window.geometry(dimensions)
        return window

    def create_scrollable_frame(self, window):
        frame = ttk.Frame(window)
        frame.pack(expand=True, fill=tk.BOTH)
        
        canvas = tk.Canvas(frame)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview, )
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

        return scrollable_frame

    def create_result_headings(self, scrollable_frame, send_email):
        headings = ["Customer Number", "Email", "Date", "Total"]
        if send_email:
            headings.append("Send")
        header_style = {"foreground": "black", "font": ("TkDefaultFont", 10, "bold")}
        for col_idx, heading in enumerate(headings):
            ttk.Label(scrollable_frame, text=heading, **header_style).grid(row=0, column=col_idx, padx=5, pady=5)

    def populate_results(self, scrollable_frame, send_email):
        self.checkboxes = []
        cell_style = {"foreground": "black"}
        for idx, item in enumerate(self.processor.final_list, start=1):
            ttk.Label(scrollable_frame, text=item['customer_num'], **cell_style).grid(row=idx, column=0, padx=5, pady=5)
            ttk.Label(scrollable_frame, text=item['email'], **cell_style).grid(row=idx, column=1, padx=5, pady=5)
            ttk.Label(scrollable_frame, text=item.get('statement_date', item.get('invoice_num', '')), **cell_style).grid(row=idx, column=2, padx=5, pady=5)
            ttk.Label(scrollable_frame, text=item['total'], **cell_style).grid(row=idx, column=3, padx=5, pady=5)

            if send_email:
                var = tk.BooleanVar(value=item['send'])
                cb = ttk.Checkbutton(scrollable_frame, variable=var, bootstyle="danger-square-toggle")
                cb.grid(row=idx, column=4, padx=5, pady=5)
                self.checkboxes.append((item, var))
            else:
                var = tk.BooleanVar(False)
                self.checkboxes.append((item, var))

    def create_unmatched_headings(self, scrollable_frame):
        headings = ["Customer Number", "Date", "Total", "Email"]
        if self.config["send_email"]:
            headings.append("Send")
        header_style = {"foreground": "black", "font": ("TkDefaultFont", 10, "bold")}
        for col_idx, heading in enumerate(headings):
            ttk.Label(scrollable_frame, text=heading, **header_style).grid(row=0, column=col_idx, padx=5, pady=5)

    def populate_unmatched(self, scrollable_frame):
        self.unmatched_vars = []
        cell_style = {"foreground": "black"}
        for idx, item in enumerate(self.processor.unmatched, start=1):
            email_var = tk.StringVar(value=item[c.EMAIL_KEY])
            send_var = tk.BooleanVar(value=item[c.SEND_KEY])
            ttk.Label(scrollable_frame, text=item['customer_num'], **cell_style).grid(row=idx, column=0, padx=5, pady=5)
            ttk.Label(scrollable_frame, text=item.get('statement_date', item.get('invoice_num', '')), **cell_style).grid(row=idx, column=1, padx=5, pady=5)
            ttk.Label(scrollable_frame, text=item['total'], **cell_style).grid(row=idx, column=2, padx=5, pady=5)
            ttk.Entry(scrollable_frame, textvariable=email_var, bootstyle="danger").grid(row=idx, column=3, sticky='ew', padx=5, pady=5)
            if self.config["send_email"]:
                ttk.Checkbutton(scrollable_frame, variable=send_var, bootstyle="danger-square-toggle").grid(row=idx, column=4, sticky='w', padx=35, pady=5)
            self.unmatched_vars.append((item, email_var, send_var))

    def create_action_buttons(self, window, command):
        action_frame = ttk.Frame(window)
        action_frame.pack(fill=tk.X, pady=10)

        action_button = ttk.Button(action_frame, text="Continue", command=lambda: command(window), bootstyle="success")
        action_button.pack(side=tk.LEFT, padx=20)

        exit_button = ttk.Button(action_frame, text="Exit", command=window.destroy, bootstyle="danger")
        exit_button.pack(side=tk.RIGHT, padx=20)

    def update_unmatched_entries(self, window):
        for item, email_var, send_var in self.unmatched_vars:
            new_email = email_var.get()
            if new_email != item[c.EMAIL_KEY]:
                item[c.EMAIL_KEY] = new_email
                self.update_email_file(item)
                self.processor.final_list.append(item)                

        window.destroy()
        self.status_label.config(text="Unmatched entries updated.")
        self.show_results()

    def update_email_file(self, item):
        try:
            df = pd.read_excel(self.config['excel'])
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel file: {self.config['excel']}")
            return

        customer_code = item['customer_num']
        email = item[c.EMAIL_KEY]

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
        root.withdraw()

        clear = messagebox.askyesno("Confirm Deletion", "Would you like to permanently delete the input document?")
        root.destroy()

        if clear:
            self.processor.clear_directory(input_dir)

    def proceed(self, window):
        if self.checkboxes:
            for item, var in self.checkboxes:
                item['send'] = var.get()

        if self.temp_dir:
            if self.config['send_email']:
                self.processor.check_and_send_documents(self.processor.final_list, self.temp_dir)
                messagebox.showinfo("Emails Sent", 'All emails flagged to send have been sent')

            if self.doc_type.get() == "invoice":
                self.processor.log_invoice_run(self.config['input'], self.temp_dir)
                self.update_invoice_text()
            else:
                self.processor.clear_directory(self.temp_dir)
            
            if self.config['clear_inputs']:
                self.processor.clear_directory(self.config['input'])
                messagebox.showinfo("Input Cleared", 'The input directory has been cleared')
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
        settings_window.geometry("550x400")

        container = ttk.Frame(settings_window)
        container.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview, bootstyle="default")
        
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
                ttk.Label(scrollable_frame, text=key, foreground="black").grid(row=row, column=0, sticky='w', pady=5, padx=10)
                var = ttk.StringVar(value=str(value))
                entry = ttk.Entry(scrollable_frame, textvariable=var, bootstyle="dark")
                entry.grid(row=row, column=1, sticky='ew', pady=5, padx= 40)
                entry.config(width=50)
                self.settings_vars[key] = var
                row += 1

        save_button = ttk.Button(settings_window, text="Save", command=self.save_settings, bootstyle="danger")
        save_button.pack(pady=10)


    def save_settings(self):
        for key, var in self.settings_vars.items():
            value = var.get()
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
        messagebox.showinfo("Settings Saved", "Configuration has been saved successfully.")

if __name__ == "__main__":
    app = DocumentProcessorApp()
    app.mainloop()
