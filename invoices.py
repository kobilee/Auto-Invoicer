import fitz
import os
import tempfile
import win32com.client as win32
import pandas as pd
from documents import DocumentProcessor

INVOICE_TOTAL = "Invoice Total"

class InvoiceProcessor(DocumentProcessor):
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

   