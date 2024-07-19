import fitz
import os
import win32com.client as win32
import pandas as pd
import shutil
import json
from src.backend.documents import DocumentProcessor
import src.backend.constants as c
INVOICE_TOTAL = "Invoice Total"
EMAIL = "email"

class InvoiceProcessor(DocumentProcessor):
    def __init__(self, config):
        super().__init__(config)
        self.doc_type = "invoice"
        
    def pdf(self, file, temp_dir):
        """
        Processes a PDF file to extract invoice and customer data,
        and saves each invoice as a separate multi-page PDF.

        Args:
            file (str): The file path for the PDF file.
        """
        with fitz.open(file) as pdf_file:
            invoices = self.extract_invoices(pdf_file)
            if invoices == "Error":
                return"Error"
            self.save_invoices_as_pdfs(invoices, pdf_file, temp_dir)
            return temp_dir

    def extract_invoices(self, pdf_file):
        """
        Extracts invoice data from each page of the PDF and groups pages by invoice number.

        Args:
            pdf_file (fitz.Document): The PDF document to process.

        Returns:
            dict: A dictionary where the keys are invoice numbers and the values are lists of page numbers.
        """
        invoices = {}
        
        for page_num, page in enumerate(pdf_file):
            invoice_data = self.extract_data_from_page(page, page_num)
            
            if not invoice_data:
                return "Error"
                
            invoice_num = invoice_data[c.FILE_KEY]
            
            if invoice_num not in invoices:
                invoices[invoice_num] = []
            
            invoices[invoice_num].append(page_num)
            self.data_list.append(invoice_data)
        
        return invoices

    def extract_data_from_page(self, page, page_num):
        """
        Extracts data from a single page of the PDF.

        Args:
            page (fitz.Page): The page to extract data from.
            page_num (int): The page number in the PDF.

        Returns:
            dict: A dictionary containing the extracted invoice data.
        """
        def extract_text_from_rect(rect):
            return page.get_text("text", clip=rect).strip()

        try:
            invoice_pos = page.search_for(self.config['pdf_invoice'])[0]
        except IndexError:
            print(f"String '{self.config['pdf_invoice']}' not found on page {page_num}")
            return None

        invoice_num_rect = fitz.Rect(invoice_pos.x1, invoice_pos.y0 - 10, invoice_pos.x1 + 150, invoice_pos.y1 + 10)
        invoice_num = extract_text_from_rect(invoice_num_rect)

        try:
            customer_pos = page.search_for(self.config['pdf_invoice_customer'])[0]
        except IndexError:
            print(f"String '{self.config['pdf_invoice_customer']}' not found on page {page_num}")
            return None

        customer_num_rect = fitz.Rect(customer_pos.x1, customer_pos.y0 - 10, customer_pos.x1 + 150, customer_pos.y1 + 10)
        customer_num = extract_text_from_rect(customer_num_rect)

        try:
            invoice_total_pos = page.search_for(c.INVOICE_TOTAL)[0]
        except IndexError:
            print(f"String '{c.INVOICE_TOTAL}' not found on page {page_num}")
            return None

        invoice_total_rect = fitz.Rect(invoice_total_pos.x1, invoice_total_pos.y0 - 10, invoice_total_pos.x1 + 150, invoice_total_pos.y1 + 10)
        invoice_total = extract_text_from_rect(invoice_total_rect)

        return {
            c.FILE_KEY: invoice_num,
            c.INVOICE_KEY: invoice_num,
            c.CUSTOMER_KEY: customer_num,
            c.TOTAL_KEY: invoice_total
        }

    def save_invoices_as_pdfs(self, invoices, pdf_file, temp_dir):
        """
        Saves grouped pages into separate PDF files for each invoice.

        Args:
            invoices (dict): A dictionary where the keys are invoice numbers and the values are lists of page numbers.
            pdf_file (fitz.Document): The original PDF document.
            temp_dir (str): The path to the temporary directory.
        """
        for invoice_num, pages in invoices.items():
            filename = invoice_num + '.pdf'
            temp_file_path = os.path.join(temp_dir, filename)

            new_pdf = fitz.open()
            for page_num in pages:
                new_pdf.insert_pdf(pdf_file, from_page=page_num, to_page=page_num)

            new_pdf.save(temp_file_path)
            new_pdf.close()


