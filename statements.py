import fitz
import os
import win32com.client as win32
import pandas as pd
from documents import DocumentProcessor

INVOICE_TOTAL = "Invoice Total"

class StatementProcessor(DocumentProcessor):
    """
    Subclass for processing statements and sending them to clients via email.
    """

    def pdf(self, file):
        """
        Reads a PDF file and extracts statement and customer data.

        Args:
            file (str): The file path for the PDF file.
        """
        with fitz.open(file) as pdf_file:
            # Custom logic for extracting statement data
            for page_num, page in enumerate(pdf_file):
                # Example: Search for the text "statement #"
                statement_pos = page.search_for(self.variable_config['pdf_statement'])[0]

                # Extract the statement number from the cell to the right
                statement_num_rect = fitz.Rect(statement_pos.x1, statement_pos.y0-10, statement_pos.x1+150, statement_pos.y1+10) # Assumes cell is 150 units wide
                statement_num = page.get_text("text", clip=statement_num_rect).strip()

                # Search for the text "Customer #"
                customer_pos = page.search_for(self.variable_config['pdf_customer'])[0]

                # Extract the customer number from the cell to the right
                customer_num_rect = fitz.Rect(customer_pos.x1, customer_pos.y0-10, customer_pos.x1+150, customer_pos.y1+10) # Assumes cell is 150 units wide
                customer_num = page.get_text("text", clip=customer_num_rect).strip()

                # Search for the text "Statement Total"
                statement_total_pos = page.search_for(INVOICE_TOTAL)[0]

                # Extract the statement total from the cell to the right
                statement_total_rect = fitz.Rect(statement_total_pos.x1, statement_total_pos.y0-10, statement_total_pos.x1+150, statement_total_pos.y1+10) # Assumes cell is 150 units wide
                statement_total = page.get_text("text", clip=statement_total_rect).strip()

                # Create a dictionary mapping "statement #" to <a statement number> and "Customer #" to <a customer number>
                data = {
                    self.variable_config['pdf_statement']: statement_num,
                    self.variable_config['pdf_customer']: customer_num,
                    INVOICE_TOTAL: statement_total
                }

                # Append the dictionary to the statement list
                self.data_list.append(data)

                # Save the current page as a PDF file with the statement number as the file name
                filename = statement_num + '.pdf'
                file_path = os.path.join(os.path.dirname(file), filename)
                # Create a new PDF document containing only the current page
                new_pdf = fitz.open()
                new_pdf.insert_pdf(pdf_file, from_page=page_num, to_page=page_num)

                new_pdf.save(file_path)

                # Close the new PDF document
                new_pdf.close()
