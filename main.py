import PyPDF2
import fitz
import os
import win32com.client
from invoices import Invoice
import setting

# path config
path_config = {
    'excel': setting.PATH_EXCEL,
    'input': setting.PATH_INPUT,
    'backup': setting.PATH_BACKUP,
}

# column data config
email_config = {
    'pdf_invoice': setting.PDF_INVOICE_STR,
    'pdf_customer': setting.PDF_CUSTOMER_STR,
    'excel_customer': setting.EXCEL_CUSTOMER_STR,
    'excel_email': setting.EXCEL_EMAIL_STR
}

option_config = {    
    'send_email': setting.SEND_EMAIL,
    'backup_pdf': setting.BACKUP_PDF,
    'cc': setting.CC
}
def main():
    invoice = Invoice(path_config['excel'], email_config, option_config)
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
            if option_config['send_email']:
                invoice.check_and_send_invoices(invoice.final_list, input_folder)
            if option_config['backup_pdf']:
                invoice.copy_and_clear_directory(input_folder, backup_folder)
    else:
        print(f"No pdf's found in the input directory: {path_config['input']} ")

    

if __name__ == '__main__':
    main()


