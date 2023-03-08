import PyPDF2
import fitz
import os
import win32com.client as win32
import pandas as pd
import shutil
import datetime

EMAIL = "email"
INVOICE_TOTAL = "Invoice Total"
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


    def __init__(self, file_path, config, options):
        self.file_path = file_path
        self.invoice_list = []
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

    def pdf(self, file):
        """
        Reads a PDF file and extracts invoice and customer data.

        Args:
            file (str): The file path for the PDF file.
        """

        with fitz.open(file) as pdf_file:
            # Loop through all the pages
            for page_num, page in enumerate(pdf_file):
                # Search for the text "invoice #"
                invoice_pos = page.search_for(self.variable_config['pdf_invoice'])[0]

                # Extract the invoice number from the cell to the right
                invoice_num_rect = fitz.Rect(invoice_pos.x1, invoice_pos.y0-10, invoice_pos.x1+150, invoice_pos.y1+10) # Assumes cell is 150 units wide
                invoice_num = page.get_text("text", clip=invoice_num_rect).strip()

                # Search for the text "Customer #"
                customer_pos = page.search_for(self.variable_config['pdf_customer'])[0]

                # Extract the customer number from the cell to the right
                customer_num_rect = fitz.Rect(customer_pos.x1, customer_pos.y0-10, customer_pos.x1+150, customer_pos.y1+10) # Assumes cell is 150 units wide
                customer_num = page.get_text("text", clip=customer_num_rect).strip()

                # Search for the text "Customer #"
                invoice_total_pos = page.search_for(INVOICE_TOTAL)[0]

                # Extract the customer number from the cell to the right
                invoice_total_rect = fitz.Rect(invoice_total_pos.x1, invoice_total_pos.y0-10, invoice_total_pos.x1+150, invoice_total_pos.y1+10) # Assumes cell is 150 units wide
                invoice_total = page.get_text("text", clip=invoice_total_rect).strip()

                # Create a dictionary mapping "invoice #" to <a invoice number> and "Customer #" to <a customer number>
                data = {
                    self.variable_config['pdf_invoice']: invoice_num,
                    self.variable_config['pdf_customer']: customer_num,
                    INVOICE_TOTAL: invoice_total
                }

                # Append the dictionary to the invoice list
                self.invoice_list.append(data)

                # Save the current page as a PDF file with the invoice number as the file name
                filename = invoice_num + '.pdf'
                file_path = os.path.join(os.path.dirname(file), filename)
                # Create a new PDF document containing only the current page
                new_pdf = fitz.open()
                new_pdf.insert_pdf(pdf_file, from_page=page_num, to_page=page_num)

                new_pdf.save(file_path)

                # Close the new PDF document
                new_pdf.close()

    def find_client(self):
        """
        Matches customer numbers from the invoice data to the client data and stores the final data in a list.
        """

        client_data_hash = {entry[self.variable_config['excel_customer']]: entry for entry in self.client_data}
        for entry in self.invoice_list:
            ccustno = entry[self.variable_config['pdf_customer']]
            if ccustno in client_data_hash:
                match = {self.variable_config['pdf_customer']: ccustno, EMAIL: client_data_hash[ccustno][self.variable_config['excel_email']], self.variable_config['pdf_invoice']: entry[self.variable_config['pdf_invoice']], INVOICE_TOTAL: entry[INVOICE_TOTAL]}
                self.final_list.append(match)

        # print the result
        print(self.final_list)

    def send_email_with_attachment(self, email_address, attachment_path, invoice_num, client_num):
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
            if self.options_config['cc']:
                mail.CC = self.options_config['cc']
            mail.Subject = f'PortaMini Invoice {invoice_num} {client_num}'
            mail.Body = '''Please see attached invoice.\n
                            \n
                            Regards,\n
                            \n
                            Rob Lee\n
                            PortaMini Storage\n
                            www.portamini.com\n
                            Phone: 416-221-6660\n'''
            
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
            invoice_number = dictionary.get(self.variable_config['pdf_invoice'])
            customer_number = dictionary.get(self.variable_config['pdf_customer'])

            invoice_path = os.path.join(directory_path, f'{invoice_number}.pdf')
            if os.path.exists(invoice_path) and dictionary.get(INVOICE_TOTAL) != "0.00":
                email_address = dictionary.get(EMAIL)
                if isinstance(email_address, str):
                    self.send_email_with_attachment(email_address, invoice_path, invoice_number, customer_number) 
                elif isinstance(email_address, list):
                    for email in email_address:
                        self.send_email_with_attachment(email, invoice_path, invoice_number, customer_number)
 
            else:
                print(f"File {invoice_path} not found for invoice {dictionary[self.variable_config['pdf_invoice']]}. Email not sent.")
       
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



