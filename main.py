import os
import argparse
from invoices import InvoiceProcessor
from statements import StatementProcessor
import setting

# Path configuration
path_config = {
    'excel': setting.PATH_EXCEL,
    'input': setting.PATH_INPUT,
    'backup': setting.PATH_BACKUP,
}

# Column data configuration
email_config = {
    'pdf_invoice': setting.PDF_INVOICE_STR,
    'pdf_customer': setting.PDF_CUSTOMER_STR,
    'excel_customer': setting.EXCEL_CUSTOMER_STR,
    'excel_email': setting.EXCEL_EMAIL_STR,
}

option_config = {    
    'send_email': setting.SEND_EMAIL,
    'backup_pdf': setting.BACKUP_PDF,
    'cc': setting.CC
}

def process_documents(processor):
    processor.read_excel_to_dict()
    input_folder = path_config["input"]
    backup_folder = path_config["backup"]

    found_pdf = False
    for filename in os.listdir(input_folder):
        if filename.endswith(".pdf"):
            file = os.path.join(input_folder, filename)
            processor.pdf(file)
            found_pdf = True
        else:
            continue
    
    if found_pdf:
        processor.find_client()

        pause = input("Please review the documents/emails in the terminal. Press 'Y' to continue or any other key to exit: ")
        if pause.upper() == "Y":
            if option_config['send_email']:
                processor.check_and_send_documents(processor.final_list, input_folder)
            if option_config['backup_pdf']:
                processor.copy_and_clear_directory(input_folder, backup_folder)
    else:
        print(f"No PDFs found in the input directory: {path_config['input']} ")

def main():
    parser = argparse.ArgumentParser(description='Process invoices or statements.')
    parser.add_argument('document_type', choices=['invoice', 'statement'], help='Type of document to process')
    args = parser.parse_args()

    if args.document_type == 'invoice':
        processor = InvoiceProcessor(path_config['excel'], email_config, option_config)
    elif args.document_type == 'statement':
        processor = StatementProcessor(path_config['excel'], email_config, option_config)

    process_documents(processor)

if __name__ == '__main__':
    main()