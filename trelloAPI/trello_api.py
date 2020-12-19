from . import card as trello_card
from . import board as trello_board
from datetime import datetime

#プロジェクトのタスク開始日を取得する
def get_project_start_date(all_board_and_card):
    if len(all_board_and_card) < 1:
        return

    all_cards = []
    for board in all_board_and_card:
        for card in board.get_cards():
            all_cards.append(card) 

    all_cards.sort(key=lambda v : v.get_date())
    return all_cards[0].get_date()

def is_task_date_in_date(card, date):
    start_date = card.get_date()
    end_date = card.get_due()
    return start_date <= date and end_date >= date

def is_task_actual_date_in_date(card, date):
    start_date = card.get_date()
    end_date = card.get_date_last_activity()
    return start_date <= date and end_date >= date

def str_to_trello_format_datetime(date_str):
    return datetime.strptime(date_str, '%Y-%m-%dT%H:%M:%S.%f%z')

class TrelloAPI:
    __cards = []

    def __init__(self, content):
        self.__board = trello_board.Board()
        self.__board.set_name(content['name'])
        self.__board.set_board_id(content['id'])
        self.__board.set_desription(content['desc'])

        listname = {}
        for v in content['lists']:
            listname[v['id']] = v['name']

        members = {}
        for v in content['members']:
            members[v['id']] = v['fullName']

        for card_json in content['cards']:
            card = trello_card.Card()
            card.set_card_id(card_json['id'])
            card.set_name(card_json['name'])
            card.set_desc(card_json['desc'])
            card.set_listname(listname[card_json['idList']])
            card.set_closed(card_json['closed'])
            card.set_date_last_activity(card_json['dateLastActivity'])
            card.set_due(card_json['due'])
            card.set_due_complete(card_json['dueComplete'])

            for action in content['actions']:
                try:
                    card_id = action['data']['card']['id']
                except KeyError:
                    continue
                if card_id == card_json['id'] :
                    card.set_date(action['date'])

            membernames = []
            for member_json in card_json['idMembers']:
                membernames.append(members[member_json])
            card.set_membernames(membernames)
            self.__cards.append(card)

        self.__board.set_cards(self.__cards)

    def get_cards(self):
        return self.__cards

    def get_board(self):
        return self.__board

    def sort_cards_by_date(self):
        self.__cards.sort(key=lambda v : v.get_date())
