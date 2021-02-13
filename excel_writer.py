from constants import IMG_PROPERTY, EXCEL_FOLDER
import xlsxwriter


class ExcelWriter:
    def __init__(self, dictionaries):
        self.route = EXCEL_FOLDER
        self.dictionaries = dictionaries

    def write_excel_files(self):
        for dictionary in self.dictionaries:
            workbook = xlsxwriter.Workbook(f"{self.route}/{dictionary['dni']}.xlsx")
            worksheet = workbook.add_worksheet()
            self.write_worksheet(worksheet, dictionary)
            workbook.close()

    def write_worksheet(self, worksheet, dictionary):
        row, col = 0, 0
        for key, value in dictionary.items():
            worksheet.write(row, col, key)
            if key == IMG_PROPERTY:
                worksheet.write_url(row, col + 1, value)
            else:
                worksheet.write(row, col + 1, value)
            row += 1
