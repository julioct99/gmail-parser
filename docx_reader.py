from constants import *
import docx2txt
import docx
import os


class DocxReader:
    def __init__(self):
        self.docs_route = DOCS_FOLDER
        self.imgs_route = IMAGES_FOLDER
        self.dictionaries = []

    def parse_docx_files(self):
        for document in os.listdir(self.docs_route):
            dictionary = self.parse_docx(docx.Document(f"{self.docs_route}/{document}"))
            self.dictionaries.append(dictionary)
        return self.dictionaries

    def parse_docx(self, document):
        paragraphs = [p.text for p in document.paragraphs]
        dictionary = self.make_dictionary(paragraphs)
        img_routes = self.extract_images(document, dictionary["dni"])
        if img_routes:
            dictionary[IMG_PROPERTY] = img_routes[0]
        return dictionary

    def extract_images(self, document, dni):
        img_routes = []
        for shape in document.inline_shapes:
            content_id = shape._inline.graphic.graphicData.pic.blipFill.blip.embed
            content_type = document.part.related_parts[content_id].content_type
            if not content_type.startswith("image"):
                continue
            img_name, img_data, extension = self.get_image(document, content_id)
            img_route = f"{self.imgs_route}/{dni}{extension}"
            img_route_excel = f"../{IMAGES_FOLDER_NAME}/{dni}{extension}"
            img_routes.append(img_route_excel)
            with open(img_route, "wb") as f:
                f.write(img_data)
        return img_routes

    def make_dictionary(self, paragraphs):
        dictionary = {}
        for paragraph in filter(lambda p: p, paragraphs):
            key, value = paragraph.split(":")
            dictionary[key.strip()] = value.strip()
        return dictionary

    def get_image(self, document, img_id):
        img_name = os.path.basename(document.part.related_parts[img_id].partname)
        img_data = document.part.related_parts[img_id]._blob
        extension = self.get_file_extension(img_name)
        return img_name, img_data, extension

    def get_file_extension(self, filename):
        _, extension = os.path.splitext(filename)
        return extension

