import logic.trello_api
import json
import csv

class ExportCsv:
    def __init__(self, json_file, csv_file):
        self.__csv_file = csv_file
        self.__json_file = json_file

    def readJson(self):
        file_open = open(self.__json_file, 'r')
        content = json.load(file_open)
        return content

    def exportAsCsv(self):
        api = logic.trello_api.TrelloAPI(self.readJson())
        file_open  = open(self.__csv_file, 'w')
        writer = csv.writer(file_open, lineterminator='\n')
        writer.writerow(['ID', 'タスク', 'タスク内容', '更新日', 'ステータス', '担当者'])
        data = []
        for v in api.get_cards() :
            content = []
            content.append(v.get_card_id())
            content.append(v.get_name())
            content.append(v.get_desc())
            content.append(v.get_date_last_activity())
            content.append(v.get_listname())
            content.append(','.join(v.get_membernames()))
            data.append(content)

        writer.writerows(data)
    
        file_open.close()
