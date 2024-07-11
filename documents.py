import PyPDF2
import fitz
import os
import win32com.client as win32
import pandas as pd
import shutil
import datetime

EMAIL = "email"
INVOICE_TOTAL = "Invoice Total"

class DocumentProcessor:
    """
    Parent class for processing documents (invoices or statements) and sending them via email.

    Attributes:
        file_path (str): The file path for the database file.
        data_list (list): A list to store the document data.
        client_data (list): A list to store the client data.
        final_list (list): A list to store the final document and client data.
        config (dict): A dictionary that stores the configuration data.
    """

    def __init__(self, file_path, config, options):
        self.file_path = file_path
        self.data_list = []
        self.client_data = []
        self.final_list = []
        self.variable_config = config
        self.options_config = options

    def read_excel_to_dict(self):
        """
        Reads the excel file and stores the client data in a list of dictionaries.
        """
        df = pd.read_excel(self.file_path)
        self.client_data = df.to_dict('records')
        
        for item in self.client_data:
            if "," in item[self.variable_config['excel_email']]:
                item[self.variable_config['excel_email']] = [x.strip() for x in item[self.variable_config['excel_email']].split(",")]

    def send_email_with_attachment(self, email_address, attachment_path, document_num, client_num):
        """
        Sends an email with an attached document PDF.

        Args:
            email_address (str): The email address of the client.
            attachment_path (str): The file path for the document PDF.
        """
        try:
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = email_address
            if self.options_config['cc']:
                mail.CC = self.options_config['cc']
            mail.Subject = f'PortaMini Document {document_num} {client_num}'
            mail.Body = '''Please see attached document.\n
                            \n
                            Regards,\n
                            \n
                            Rob Lee\n
                            PortaMini Storage\n
                            www.portamini.com\n
                            Phone: 416-221-6660\n'''
            
            mail.Attachments.Add(attachment_path)
            mail.Send()
            print(f"Email sent to {email_address} with document {attachment_path}.")
        except Exception as e:
            print(f"Email to {email_address} failed to send. Please send document {attachment_path} manually. Error: {e}")

    def find_client(self):
        """
        Matches customer numbers from the document data to the client data and stores the final data in a list.
        """
        client_data_hash = {entry[self.variable_config['excel_customer']]: entry for entry in self.client_data}
        for entry in self.data_list:
            ccustno = entry[self.variable_config['pdf_customer']]
            if ccustno in client_data_hash:
                match = {
                    self.variable_config['pdf_customer']: ccustno,
                    EMAIL: client_data_hash[ccustno][self.variable_config['excel_email']],
                    self.variable_config['pdf_invoice']: entry[self.variable_config['pdf_invoice']],
                    INVOICE_TOTAL: entry[INVOICE_TOTAL]
                }
                self.final_list.append(match)

        print(self.final_list)

    def check_and_send_documents(self, dicts_list, directory_path):
        """
        Checks if the document PDF exists and sends an email with the PDF attached if it does.

        Args:
            dicts_list (list): A list of dictionaries containing document and client data.
            directory_path (str): The directory path where the document PDFs are stored.
        """
        for dictionary in dicts_list:
            document_number = dictionary.get(self.variable_config['pdf_invoice'])
            customer_number = dictionary.get(self.variable_config['pdf_customer'])

            document_path = os.path.join(directory_path, f'{document_number}.pdf')
            if os.path.exists(document_path) and dictionary.get(INVOICE_TOTAL) != "0.00":
                email_address = dictionary.get(EMAIL)
                if isinstance(email_address, str):
                    self.send_email_with_attachment(email_address, document_path, document_number, customer_number) 
                elif isinstance(email_address, list):
                    for email in email_address:
                        self.send_email_with_attachment(email, document_path, document_number, customer_number)
            else:
                print(f"File {document_path} not found for document {dictionary[self.variable_config['pdf_invoice']]}. Email not sent.")

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
        print(f'The input directory has been cleared and today\'s documents have been backed up into: {new_dir}')