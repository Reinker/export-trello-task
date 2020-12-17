import trelloAPI.trello_api as trello_api
import trelloAPI.card as trello_card
from openpyxl import Workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.styles import PatternFill, Border, Side
import calendar
from datetime import datetime
import json
import csv
import os

ROW_START=1
NORMAL_SIDE = Side(style='thin', color='000000')
BORDER = Border(top=NORMAL_SIDE,bottom=NORMAL_SIDE,right=NORMAL_SIDE,left=NORMAL_SIDE)
TOP_FILL = PatternFill(fill_type='solid', fgColor='55FF55')
PHASE_FILL = PatternFill(fill_type='solid', fgColor='FFFF55')

class ExportExcel:

    def __init__(self, json_file):
        self.__json_file = json_file
        self.__boards = []
        self.__wb = Workbook()
        self.__ws = self.__wb.active
        self.__ws.title = 'タスク一覧'
        self.__project_start_date = ''

    def import_from_files(self):
        files = os.listdir('./jsons')
        if files.count == 0 :
            return

        for f in files:
            print('load file : ' + f)
            file_open = open('./jsons/' + f, 'r')
            json_str = json.load(file_open)
            api = trello_api.TrelloAPI(json_str)
            self.__boards.append(api.get_board())
            api.sort_cards_by_date()

        #self.__project_start_date = datetime.strptime(trello_api.get_project_start_date(self.__boards), '%Y-%m-%dT%H:%M:%S.%f%z')

        #months = []
        #months.append(calendar.monthcalendar(self.__project_start_date.year, self.__project_start_date.month))
        #try:
        #    months.append(calendar.monthcalendar(self.__project_start_date.year, self.__project_start_date.month + 1))
        #except calendar.IllegalMonthError:
        #    months.append(calendar.monthcalendar(self.__project_start_date.year, calendar.January))

        #try:
        #    months.append(calendar.monthcalendar(self.__project_start_date.year, self.__project_start_date.month + 2))
        #except calendar.IllegalMonthError:
        #    months.append(calendar.monthcalendar(self.__project_start_date.year, calendar.February))


    def performance(self):
        if len(self.__boards) < 1:
            return

        for board in self.__boards:
            for card in board.get_cards():
                print(card.get_date())
    
    def __set_item_name_cell(self, col_num, width, name):
        self.__ws.column_dimensions[get_column_letter(col_num)].width = width
        self.__ws.cell(row=ROW_START, column=col_num).value = name
        self.__ws.cell(row=ROW_START, column=col_num).fill = TOP_FILL
        self.__ws.cell(row=ROW_START, column=col_num).border = BORDER

    def task_ids(self, col_num): 
        self.__set_item_name_cell(col_num, 30, 'ID')

        row = ROW_START + 1
        for board in self.__boards:
            self.__ws.cell(row=row, column=col_num).value = board.get_board_id()
            self.__ws.cell(row=row, column=col_num).fill = PHASE_FILL
            self.__ws.cell(row=row, column=col_num).border = BORDER 
            row += 1
            for card in board.get_cards():
                self.__ws.cell(row=row, column=col_num).value = card.get_card_id() 
                row += 1

    def task_names(self, col_num):
        self.__set_item_name_cell(col_num, 30, 'タスク名')
        row = ROW_START + 1
        for board in self.__boards:
            self.__ws.cell(row=row, column=col_num).value = board.get_name()
            self.__ws.cell(row=row, column=col_num).fill = PHASE_FILL
            self.__ws.cell(row=row, column=col_num).border = BORDER 
            row += 1
            for card in board.get_cards():
                self.__ws.cell(row=row, column=col_num).value = card.get_name()
                row += 1
    
    def task_description(self, col_num):
        self.__set_item_name_cell(col_num, 40, '説明')
        row = ROW_START + 1
        for board in self.__boards:
            self.__ws.cell(row=row, column=col_num).value = board.get_description()
            self.__ws.cell(row=row, column=col_num).fill = PHASE_FILL
            self.__ws.cell(row=row, column=col_num).border = BORDER 
            row += 1
            for card in board.get_cards():
                self.__ws.cell(row=row, column=col_num).value = card.get_desc()
                row += 1

    def task_start_date(self, col_num):
        self.__set_item_name_cell(col_num, 30, '開始日')
        row = ROW_START + 1
        for board in self.__boards:
            self.__ws.cell(row=row, column=col_num).fill = PHASE_FILL
            self.__ws.cell(row=row, column=col_num).border = BORDER 
            row += 1
            for card in board.get_cards():
                self.__ws.cell(row=row, column=col_num).value = card.get_date()
                row += 1

    def task_last_activity_date(self, col_num):
        self.__set_item_name_cell(col_num, 30, '更新日')
        row = ROW_START + 1
        for board in self.__boards:
            self.__ws.cell(row=row, column=col_num).fill = PHASE_FILL
            self.__ws.cell(row=row, column=col_num).border = BORDER 
            row += 1
            for card in board.get_cards():
                self.__ws.cell(row=row, column=col_num).value = card.get_date_last_activity()
                row += 1

    def task_list_name(self, col_num):
        self.__set_item_name_cell(col_num, 20, 'ステータス')
        row = ROW_START + 1
        for board in self.__boards:
            self.__ws.cell(row=row, column=col_num).fill = PHASE_FILL
            self.__ws.cell(row=row, column=col_num).border = BORDER 
            row += 1
            for card in board.get_cards():
                self.__ws.cell(row=row, column=col_num).value = card.get_listname()
                row += 1

    def task_members(self, col_num):
        self.__set_item_name_cell(col_num, 20, '担当者')
        row = ROW_START + 1
        for board in self.__boards:
            self.__ws.cell(row=row, column=col_num).fill = PHASE_FILL
            self.__ws.cell(row=row, column=col_num).border = BORDER 
            row += 1
            for card in board.get_cards():
                self.__ws.cell(row=row, column=col_num).value = ','.join(card.get_membernames())
                row += 1


    def exportAsExcel(self):
        if len(self.__boards) < 1:
            return

        self.task_ids(1) 
        self.task_names(2)
        self.task_description(3)
        self.task_start_date(4)
        self.task_last_activity_date(5)
        self.task_list_name(6)
        self.task_members(7)

        self.__wb.save('test.xlsx')
