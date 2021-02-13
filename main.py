from constants import *
from gmail_parser import GmailParser
from docx_reader import DocxReader
from excel_writer import ExcelWriter
import os

if not os.path.exists(ATTACHMENTS_FOLDER):
    print(f"Creating '{ATTACHMENTS_FOLDER}' directory to store attachments...")
    os.makedirs(ATTACHMENTS_FOLDER)
    os.makedirs(DOCS_FOLDER)
    os.makedirs(EXCEL_FOLDER)
    os.makedirs(IMAGES_FOLDER)

print('\n*** [PARSING EMAILS] ***')
email_parser = GmailParser()
email_parser.parse_emails()

print(f"\n\n*** [READING DOCUMENTS IN {DOCS_FOLDER}] ***")
docx_reader = DocxReader(DOCS_FOLDER, IMAGES_FOLDER)
dictionaries = docx_reader.parse_docx_files()
print(f"{len(dictionaries)} {'document' if len(dictionaries) == 1 else 'documents'} found.")

print(f"\n\n*** [WRITING EXCEL FILES IN {EXCEL_FOLDER}] ***")
excel_writer = ExcelWriter(EXCEL_FOLDER, dictionaries)
excel_writer.write_excel_files()
