from constants import *
from gmail_parser import GmailParser
from docx_reader import DocxReader
from excel_writer import ExcelWriter
import os
import shutil

if not os.path.exists(ATTACHMENTS_FOLDER):
    print(f"Creating '{ATTACHMENTS_FOLDER}' directory to store attachments...")
    os.makedirs(ATTACHMENTS_FOLDER)
    os.makedirs(DOCS_FOLDER)
    os.makedirs(EXCEL_FOLDER)
    os.makedirs(IMAGES_FOLDER)
    
print("\n*** [PARSING EMAILS] ***")
email_parser = GmailParser("juliotestemail00@gmail.com")
email_parser.parse_emails()

print(f"\n\n*** [READING DOCUMENTS AND STORING THEM IN {DOCS_FOLDER} ***")
docx_reader = DocxReader()
dictionaries = docx_reader.parse_docx_files()
print(f"{len(dictionaries)} document{'s' if len(dictionaries) > 1 else ''} found.")

print(f"\n\n*** [WRITING EXCEL FILES IN {EXCEL_FOLDER}] ***")
excel_writer = ExcelWriter(dictionaries)
excel_writer.write_excel_files()

shutil.rmtree(TMP_FOLDER, ignore_errors=True)
