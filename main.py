import numpy as np
import math
import pandas as pd
import PyPDF2
import fitz
import os
import win32com.client
import dbf
from invoices import Invoice
def main():
    invoice = Invoice("C:/Users/kalee/projects/Auto-Invoicer/database/emailfile.dbf")
    invoice.read_dbf_to_dict()
    input_folder = "C:/Users/kalee/projects/Auto-Invoicer/pdf"
    backup_folder = "C:/Users/kalee/projects/Auto-Invoicer/backup"

    for filename in os.listdir(input_folder):
        if filename.endswith(".pdf"):
            # Do something with the file
            file = os.path.join(input_folder, filename)
            invoice.pdf(file)
        else:
            continue
    invoice.find_client()

    pause = input("Please review the invoices/emails in the termial. Press 'Y' to continue or any other key to exit: ")
    if pause.upper() == "Y":
        # invoice.check_and_send_invoices(invoice.final_list, input_folder)
        invoice.copy_and_clear_directory(input_folder, backup_folder)

    

if __name__ == '__main__':
    main()


