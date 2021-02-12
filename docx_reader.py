from constants import *
import docx2txt
import docx
import os
 
class DocxReader:
    def __init__(self, route):
        self.route = route
        self.dictionaries = []

    def parse_docx_files(self):
        for document in os.listdir(self.route):
            dict = self.parse_docx(docx.Document(f"{self.route}/{document}"))
            self.dictionaries.append(dict)
        return self.dictionaries

    def parse_docx(self, document):
        paragraphs = [p.text for p in document.paragraphs]
        return self.make_dictionary(paragraphs)

    def make_dictionary(self, paragraphs):
        dict = {}
        for paragraph in paragraphs:
            key, value = paragraph.split(':')
            dict[key.strip()] = value.strip()
        return dict

if __name__ == '__main__':
    dx_reader = DocxReader(DOCS_FOLDER)
    dictionaries = dx_reader.parse_docx_files()
    print(dictionaries)