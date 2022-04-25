from abc import ABCMeta, abstractstaticmethod
from typing import List

import openpyxl
import xlsxwriter


class RegistryBase(ABCMeta):

    REGISTRY = {}

    def __new__(cls, cls_name, bases, attrs):
        created_class = super().__new__(cls, cls_name, bases, attrs)
        if bases:  # BaseDateProvider не должен регистрироваться
            cls.REGISTRY[attrs["file_format"]] = created_class
        return created_class

    @classmethod
    def get_registry(cls):
        return dict(cls.REGISTRY)


class BaseDateProvider(metaclass=RegistryBase):

    @abstractstaticmethod
    def generate_file(filename: str, categories: List[dict]) -> None:
        pass

    @abstractstaticmethod
    def parse_file(filename: str) -> List:
        pass

class XlsxDataProvider(BaseDateProvider):
    file_format = "xlsx"

    @staticmethod
    def generate_file(filename: str, categories: List[dict]) -> None:
        workbook = xlsxwriter.Workbook(filename)
        ws = workbook.add_worksheet()
        ws.write(0, 0, "Category")
        ws.write(0, 1, "Group")
        ws.write(0, 2, "Name")
        ws.write(0, 3, "Explicits")
        ws.write(0, 4, "Implicits")
        ws.write(0, 5, "Mean")

        row = 1
        for category in categories:
            column = 0
            for item in category:
                explicits = ""
                implicits = ""
                if item.get('explicits'):
                    for i in item['explicits']:
                        explicits += "{}&".format(i)
                if item.get('implicits'):
                    for i in item['implicits']:
                        implicits += "{}&".format(i)

                ws.write(row, column, item['category'])
                ws.write(row, column + 1, item['group'])
                ws.write(row, column + 2, item['name'])
                ws.write(row, column + 3, explicits)
                ws.write(row, column + 4, implicits)
                ws.write(row, column + 5, item['mean'])
                row += 1
            row += 1

        workbook.close()

    @staticmethod
    def parse_file(filename: str) -> List:
        workbook = openpyxl.load_workbook(filename)
        sheet = workbook.active

        parsed_items_data = []
        for row in sheet.iter_rows(min_row=2, max_col=6):
            parsed_items_data.append({
                "category": row[0].value,
                "group": row[1].value,
                "name": row[2].value,
                "explicits": row[3].value,
                "implicits": row[4].value,
                "mean": row[5].value
            })

        return parsed_items_data
