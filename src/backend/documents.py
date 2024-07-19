import os
import json
import win32com.client as win32
import pandas as pd
import shutil
import datetime
import src.backend.constants as c

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

    def __init__(self, config):
        self.file_path = config["excel"]
        self.data_list = []
        self.client_data = []
        self.final_list = []
        self.unmatched = []
        self.config = config
        self.doc_type = ""

    def read_excel_to_dict(self):
        """
        Reads the excel file and stores the client data in a list of dictionaries.
        """
        df = pd.read_excel(self.file_path)
        self.client_data = df.to_dict('records')
        
        for item in self.client_data:
            if "," in item[self.config['excel_email']]:
                item[self.config['excel_email']] = [x.strip() for x in item[self.config['excel_email']].split(",")]

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
            if self.config['cc']:
                mail.CC = self.config['cc']
            mail.Subject = f'PortaMini {self.doc_type} {document_num if self.doc_type == "invoice" else ""} {client_num}'
            mail.Body = f'''Please see attached {self.doc_type}.\n
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


    def check_and_send_documents(self, dicts_list, directory_path):
        """
        Checks if the document PDF exists and sends an email with the PDF attached if it does.

        Args:
            dicts_list (list): A list of dictionaries containing document and client data.
            directory_path (str): The directory path where the document PDFs are stored.
        """
        for dictionary in dicts_list:
            document_number = dictionary.get(c.FILE_KEY)
            customer_number = dictionary.get(c.CUSTOMER_KEY)

            document_path = os.path.join(directory_path, f'{document_number}.pdf')
    
            if os.path.exists(document_path):
                if not dictionary.get(c.SEND_KEY):
                    print(f'email NOT sent to {customer_number}')
                else:
                    email_address = dictionary.get(c.EMAIL_KEY)
                    if isinstance(email_address, str):
                        self.send_email_with_attachment(email_address, document_path, document_number, customer_number) 
                    elif isinstance(email_address, list):
                        for email in email_address:

                            self.send_email_with_attachment(email, document_path, document_number, customer_number)
            else:
                print(f"File {document_path} not found for document {document_number}. Email not sent.")

    def copy_directory_with_timestamp(self, source_dir, dest_dir, doc_type):
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
        new_dir = os.path.join(dest_dir, doc_type + '_' + timestamp)
        if not os.path.isdir(new_dir):
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

    def log_invoice_run(self, input_path, temp_dir):
        """
        Copies the contents of a directory to a new directory with a timestamp, clears the original directory, and returns the
        name of the new directory.

        Args:
            source_dir (str): The path of the directory to be copied.
            dest_dir (str): The path of the directory where the copied directory with timestamp will be created.
        """
        log_file = "logs/invoice_runs.log"
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = (
            f"{timestamp} - "
            f"send_email: {self.config['send_email']} - "
            f"clear_inputs: {self.config['clear_inputs']} - "
            f"cc: {self.config['cc']} - "
            f"{', '.join(os.listdir(input_path))}\n"
        )
        with open(log_file, "a") as file:
            file.write(log_entry)
            
        self.clear_directory(temp_dir)
        print(f'The Temp directory has been cleared')
    
    def find_client(self):
        """
        Matches customer numbers from the document data to the client data and stores the final data in a list.
        """

        client_data_hash = {entry[self.config['excel_customer']]: entry for entry in self.client_data}
        for entry in self.data_list:
            ccustno = entry[c.CUSTOMER_KEY]
            if ccustno in client_data_hash:
                match = entry
                match[c.EMAIL_KEY] = client_data_hash[ccustno][self.config['excel_email']]
                total_str = match[c.TOTAL_KEY]
                total_float = float(total_str.replace(",", ""))
                match[c.SEND_KEY] = True if total_float > 0 else False
                self.final_list.append(match)
            else:
                unmatched = match = entry
                match[c.EMAIL_KEY] = ""
                match[c.SEND_KEY] = False
                self.unmatched.append(unmatched)
        
        print("The following documents will be sent:")
        for item in self.final_list:
            print(json.dumps(item, indent=4))

        print("\nThe following documents will not be sent due to Customer codes not have an email in the emailfile:")
        for item in self.unmatched:
            print(json.dumps(item, indent=4))