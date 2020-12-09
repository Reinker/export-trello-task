import logic.trello_api
from openpyxl import Workbook
from openpyxl.utils import column_index_from_string
from openpyxl.styles import PatternFill
import json
import csv
import os

class ExportExcel:
    def __init__(self, json_file):
        self.__json_file = json_file
        self.__contents = {}
        self.__tags = ['ID', 'タスク名', 'タスク説明', '更新日', 'ステータス', '担当者']
        self.__wb = Workbook()
        self.__ws = self.__wb.active
        self.__ws.title = 'タスク一覧'
        self.__fill = PatternFill(fill_type='solid', fgColor='00FF00'), '担当者'

    def __import_from_file(self):
        files = os.listdir("./jsons")
        for f in files:
            file_open = open(f, 'r')
            content = json.load(file_open)

            self.__contents.append(content)

    def __readJson(self):
        file_open = open(self.__json_file, 'r')
        content = json.load(file_open)
        return content

    def __setTags(self):
        self.__ws.column_dimensions['A'].width = 30
        for row in self.__ws.iter_rows(min_row=1, max_col=len(self.__tags), max_row=1):
            for cell in row:
                cell.fill = self.__fill
                cell.value = self.__tags[cell.column - 1]

    def exportAsExcel(self):
        self.__setTags()
        api = logic.trello_api.TrelloAPI(self.__readJson())

        cards = api.get_cards()
        for row in range(2, len(cards) + 1):
            content = []
            content.append(cards[row - 2].get_card_id())
            content.append(cards[row - 2].get_name())
            content.append(cards[row - 2].get_desc())
            content.append(cards[row - 2].get_date_last_activity())
            content.append(cards[row - 2].get_listname())
            content.append(','.join(cards[row - 2].get_membernames()))
            for col in range(1, len(self.__tags)):
                self.__ws.cell(row=row, column=col).value = content[col - 1]
        self.__wb.save('test.xlsx')
