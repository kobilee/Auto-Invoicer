import PyPDF2
import fitz
import os
import win32com.client as win32
import dbf
import shutil
import datetime

EMAIL = "email"
class Invoice:
    """
    Class for reading invoices and sending them to clients via email.

    Attributes:
        file_path (str): The file path for the database file.
        invoice_list (list): A list to store the invoice data.
        client_data (list): A list to store the client data.
        final_list (list): A list to store the final invoice and client data.
        config (dict): A dictionary that stores the configuration data.
    """


    def __init__(self, file_path, config):
        self.file_path = file_path
        self.invoice_list = []
        self.client_data = []
        self.final_list = []
        self.config = config

        
    def read_dbf_to_dict(self):
        """
        Reads the database file and stores the client data in a list of dictionaries.
        """
        table = dbf.Table(self.file_path)
        table.open()
        self.client_data = [dict((k, v.strip()) if isinstance(v, str) else (k, v) for k, v in zip(table.field_names, row)) for row in table]
        table.close()

    def pdf(self, file):
        """
        Reads a PDF file and extracts invoice and customer data.

        Args:
            file (str): The file path for the PDF file.
        """
        with fitz.open(file) as pdf_file:
            # Loop through all the pages
            for page in pdf_file:
                # Search for the text "invoice #"
                invoice_pos = page.search_for(self.config['pdf_invoice'])[0]

                # Extract the invoice number from the cell to the right
                invoice_num_rect = fitz.Rect(invoice_pos.x1, invoice_pos.y0-10, invoice_pos.x1+150, invoice_pos.y1+10) # Assumes cell is 150 units wide
                invoice_num = page.get_text("text", clip=invoice_num_rect).strip()
                
                # Search for the text "Customer #"
                customer_pos = page.search_for(self.config['pdf_customer'])[0]

                # Extract the customer number from the cell to the right
                customer_num_rect = fitz.Rect(customer_pos.x1, customer_pos.y0-10, customer_pos.x1+150, customer_pos.y1+10) # Assumes cell is 150 units wide
                customer_num = page.get_text("text", clip=customer_num_rect).strip()

                # Create a dictionary mapping "invoice #" to <a invoice number> and "Customer #" to <a customer number>
                data = {
                    self.config['pdf_invoice']: invoice_num,
                    self.config['pdf_customer']: customer_num
                }
                
                self.invoice_list.append(data)

    def find_client(self):
        """
        Matches customer numbers from the invoice data to the client data and stores the final data in a list.
        """
        self.client_data[0][self.config['dbf_customer']] = self.invoice_list[0]["Customer #"]

        client_data_hash = {entry[self.config['dbf_customer']]: entry for entry in self.client_data}

        for entry in self.invoice_list:
            ccustno = entry[self.config['pdf_customer']]
            if ccustno in client_data_hash:
                match = {self.config['pdf_customer']: ccustno, EMAIL: "kobiatlaslee@mail.com", self.config['pdf_invoice']: entry[self.config['pdf_invoice']]}
                # match = {self.config['pdf_customer']: ccustno, EMAIL: client_data_hash[ccustno][self.config['dbf_email']], self.config['pdf_invoice']: entry[self.config['pdf_invoice']]}
                self.final_list.append(match)

        # print the result
        print(self.final_list)

    def send_email_with_attachment(self, email_address, attachment_path):
        """
        Sends an email with an attached invoice PDF.

        Args:
            email_address (str): The email address of the client.
            attachment_path (str): The file path for the invoice PDF.
        """
        try:
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = email_address
            mail.Subject = 'Invoice Attachment'
            mail.Body = 'Please find attached the invoice you requested.'
            mail.Attachments.Add(attachment_path)
            mail.Send()
            print(f"Email sent to {email_address} with invoice {attachment_path}.")
        except:
            print(f"Email to {email_address} fail to send, pleae send invoice {attachment_path} maunually")

    def check_and_send_invoices(self, dicts_list, directory_path):
        """
        Checks if the invoice PDF exists and sends an email with the PDF attached if it does.

        Args:
            dicts_list (list): A list of dictionaries containing invoice and client data.
            directory_path (str): The directory path where the invoice PDFs are stored.
        """
        for dictionary in dicts_list:
            invoice_number = dictionary.get(self.config['pdf_invoice'])
            invoice_path = os.path.join(directory_path, f'{invoice_number}.pdf')
            if os.path.exists(invoice_path):
                email_address = dictionary.get(EMAIL)
                self.send_email_with_attachment(email_address, invoice_path)   
            else:
                print(f"File {invoice_path} not found for invoice {dictionary[self.config['pdf_invoice']]}. Email not sent.")
       
    def copy_directory_with_timestamp(self, source_dir, dest_dir):
        """
        Copies the contents of a directory to a new directory with a timestamp and returns the new directory name.

        Args:
        source_dir (str): The path of the directory to be copied.
        dest_dir (str): The path of the directory where the copied directory with timestamp will be created.

        Returns:
        str: The name of the new directory created with timestamp.
        """
        now = datetime.datetime.now()
        timestamp = now.strftime('%Y-%m-%d')
        new_dir = os.path.join(dest_dir, os.path.basename(source_dir) + '_' + timestamp)
        shutil.copytree(source_dir, new_dir)
        return new_dir

    def clear_directory(self, directory):
        """
        Clears the contents of a directory.

        Args:
        directory (str): The path of the directory to be cleared.
        """
        for filename in os.listdir(directory):
            file_path = os.path.join(directory, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print(f"Failed to delete {file_path}. Reason: {e}")

    def copy_and_clear_directory(self, source_dir, dest_dir):
        """
        Copies the contents of a directory to a new directory with a timestamp, clears the original directory, and returns the
        name of the new directory.

        Args:
        source_dir (str): The path of the directory to be copied.
        dest_dir (str): The path of the directory where the copied directory with timestamp will be created.

        """
        new_dir = self.copy_directory_with_timestamp(source_dir, dest_dir)
        self.clear_directory(source_dir)
        print(f'The input directory has been clear and todays invoices have been backup into: {new_dir}')



