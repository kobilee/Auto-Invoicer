import fitz
import os
import win32com.client as win32
import pandas as pd
import shutil
import json
from src.backend.documents import DocumentProcessor
import src.backend.constants as c


class StatementProcessor(DocumentProcessor):
    def pdf(self, file, temp_dir):
        """
        Processes a PDF file to extract statement and customer data,
        and saves each statement as a separate multi-page PDF.

        Args:
            file (str): The file path for the PDF file.
        """
        with fitz.open(file) as pdf_file:
            statements = self.extract_statements(pdf_file)
            self.save_statements_as_pdfs(statements, pdf_file, temp_dir)
            print(temp_dir)
            return temp_dir

    def extract_statements(self, pdf_file):
        """
        Extracts statement data from each page of the PDF and groups pages by statement number.

        Args:
            pdf_file (fitz.Document): The PDF document to process.

        Returns:
            dict: A dictionary where the keys are statement numbers and the values are lists of page numbers.
        """
        statements = {}
        for page_num, page in enumerate(pdf_file):
            statement_data = self.extract_data_from_page(page, page_num)
            if statement_data:
                statement_file = statement_data[c.FILE_KEY]
                if statement_file not in statements:
                    statements[statement_file] = []
                statements[statement_file].append(page_num)
                self.data_list.append(statement_data)
        return statements

    def extract_data_from_page(self, page, page_num):
        """
        Extracts data from a single page of the PDF.

        Args:
            page (fitz.Page): The page to extract data from.
            page_num (int): The page number in the PDF.

        Returns:
            dict: A dictionary containing the extracted statement data.
        """
        def extract_text_from_rect(rect):
            return page.get_text("text", clip=rect).strip()

        try:
            statement_pos = page.search_for(self.config['pdf_statement'])[0]
        except IndexError:
            print(f"String '{self.config['pdf_statement']}' not found on page {page_num}")
            return None

        statement_date_rect = fitz.Rect(statement_pos.x0 - 10, statement_pos.y0 + 10, statement_pos.x1 + 20, statement_pos.y1 + 20)
        statement_date = extract_text_from_rect(statement_date_rect)
        
        try:
            customer_pos = page.search_for(self.config['pdf_statement_customer'])[0]
        except IndexError:
            print(f"String '{self.config['pdf_statement_customer']}' not found on page {page_num}")
            return None

        customer_num_rect = fitz.Rect(customer_pos.x0 - 10, customer_pos.y0 + 10, customer_pos.x1 + 20, customer_pos.y1 + 40)
        customer_num = extract_text_from_rect(customer_num_rect)

        # Initialize statement_total to None
        statement_total = None
        try:
            statement_total_pos = page.search_for(c.STATEMENT_TOTAL)[0]
            statement_total_rect = fitz.Rect(statement_total_pos.x0 - 10, statement_total_pos.y0 + 10, statement_total_pos.x1 + 40, statement_total_pos.y1 + 15)
            statement_total = extract_text_from_rect(statement_total_rect)
        except IndexError:
            pass

        try:
            page_pos = page.search_for("page")[0]
            page_number_rect = fitz.Rect(page_pos.x1, page_pos.y0, page_pos.x1 + 5, page_pos.y1)
            page_number = extract_text_from_rect(page_number_rect)
        except IndexError:
            page_number = "1"  # Default to 1 if no page number is found

        
        if statement_total:
            data = {
                c.FILE_KEY: statement_date.replace('/', '_') + "_" + customer_num,
                c.STATEMENT_KEY: statement_date,
                c.CUSTOMER_KEY: customer_num,
                c.PAGE_KEY: page_number,
                c.TOTAL_KEY: statement_total
            }
            return data
        return
        

    def save_statements_as_pdfs(self, statements, pdf_file, temp_dir):
        """
        Saves grouped pages into separate PDF files for each statement.

        Args:
            statements (dict): A dictionary where the keys are statement numbers and the values are lists of page numbers.
            pdf_file (fitz.Document): The original PDF document.
            temp_dir (str): The path to the temporary directory.
        """
        for statement_file, pages in statements.items():
            filename = statement_file + '.pdf'
            temp_file_path = os.path.join(temp_dir, filename)

            new_pdf = fitz.open()
            for page_num in pages:
                new_pdf.insert_pdf(pdf_file, from_page=page_num, to_page=page_num)

            new_pdf.save(temp_file_path)
            new_pdf.close()

