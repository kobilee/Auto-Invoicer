import os
import argparse
import json
import tempfile
from src.backend.invoices import InvoiceProcessor
from src.backend.statements import StatementProcessor

def load_settings(json_path):
    if not os.path.isfile(json_path):
        raise FileNotFoundError(f"The configuration file '{json_path}' was not found.")
    
    with open(json_path, 'r') as file:
        return json.load(file)

def process_documents(processor, config, args):
    processor.read_excel_to_dict()
    input_folder = config["input"]
    backup_folder = config["backup"]
    temp_dir = tempfile.mkdtemp()

    found_pdf = False
    for filename in os.listdir(input_folder):
        if filename.endswith(".pdf"):
            file = os.path.join(input_folder, filename)
            processor.pdf(file, temp_dir)
            found_pdf = True
        else:
            continue


    if found_pdf:
        processor.find_client()

        pause = input("Please review the documents/emails in the terminal. Press 'Y' to continue or any other key to exit: ")
        if pause.upper() == "Y":
            if config['send_email']:
                processor.check_and_send_documents(processor.final_list, temp_dir)
            if config['backup_pdf']:
                processor.copy_and_clear_directory(temp_dir, backup_folder, args.document_type)
    else:
        print(f"No PDFs found in the input directory: {config['input']} ")
    print("Complete")

def main():
    parser = argparse.ArgumentParser(description='Process invoices or statements.')
    parser.add_argument('document_type', choices=['invoice', 'statement'], help='Type of document to process')
    parser.add_argument('--config', default='src/config/setting.json', help='Path to the configuration JSON file')
    args = parser.parse_args()

    config = load_settings(args.config)

    if args.document_type == 'invoice':
        processor = InvoiceProcessor(config)
    elif args.document_type == 'statement':
        processor = StatementProcessor(config)

    process_documents(processor, config, args)

if __name__ == '__main__':
    main()