import PyPDF2
import fitz
import os
import win32com.client
import dbf
from invoices import Invoice
import setting

# path config
path_config = {
    'dbf': setting.PATH_DBF,
    'input': setting.PATH_INPUT,
    'backup': setting.PATH_BACKUP,
    'send_email': setting.SEND_EMAIL,
    'backup_pdf': setting.BACKUP_PDF
}

# column data config
email_config = {
    'pdf_invoice': setting.PDF_INVOICE_STR,
    'pdf_customer': setting.PDF_CUSTOMER_STR,
    'dbf_customer': setting.DBF_CUSTOMER_STR,
    'dbf_email': setting.DBF_EMAIL_STR
}
def main():
    invoice = Invoice(path_config['dbf'], email_config)
    invoice.read_excel_to_dict()
    input_folder = path_config["input"]
    backup_folder = path_config["backup"]

    found_pdf = False
    for filename in os.listdir(input_folder):
        if filename.endswith(".pdf"):
            # Do something with the file
            file = os.path.join(input_folder, filename)
            invoice.pdf(file)
            found_pdf = True
        else:
            continue
    
    if found_pdf:
        invoice.find_client()

        pause = input("Please review the invoices/emails in the termial. Press 'Y' to continue or any other key to exit: ")
        if pause.upper() == "Y":
            if path_config['send_email']:
                invoice.check_and_send_invoices(invoice.final_list, input_folder)
            if path_config['backup_pdf']:
                invoice.copy_and_clear_directory(input_folder, backup_folder)

    

if __name__ == '__main__':
    main()


