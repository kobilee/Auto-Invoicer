## Auto-Invoicer

This project is a tool that helps automate the invoicing process. It allows the user to send invoices to clients via email and backup the invoices once sent.

# Installation

To install the Auto-Invoicer tool, follow these steps:

1. Clone the repository to your local machine.
2. Install the required dependencies by running the command `pip install -r requirements.txt`.
3. Update the configuration in `settings.py` to match your environment.

# Usage

To use the Auto-Invoicer tool, follow these steps:

1. Open a terminal.
2. Navigate to the Auto-Invoicer directory.
3. Update the `settings.py` file to point to valid paths in your system and with the correct column names. Here are the configuration options you need to update:
   - PATH_DBF: The full path to where the dbf file is located. This file must contain a column with the client number and email.
   - PATH_INPUT: The path where all the pdfs live.
   - PATH_BACKUP: The path to backup invoice once emails are sent.
   - SEND_EMAIL: A flag to enable sending emails.
   - BACKUP_PDF: A flag to enable backing up pdfs.
   - PDF_INVOICE_STR: The string in the pdf directly to the left of the invoice number.
   - PDF_CUSTOMER_STR: The string in the pdf directly to the left of the customer number.
   - DBF_EMAIL_STR: The name of the column in the dbf file that contains the email address.
   - DBF_CUSTOMER_STR: The name of the column in the dbf file that contains the customer number.
4. Run the command `python -m main` to execute the tool.
