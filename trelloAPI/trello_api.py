from trelloAPI import card as trello_card
from trelloAPI import board as trello_board
from trelloAPI import check_lists as trello_check_lists
from datetime import datetime
from datetime import date
from datetime import timezone
import calendar

TRELLO_DATETIME_FORMAT = '%Y-%m-%dT%H:%M:%S.%f%z'
INVALID_DATE_TIME = datetime.strptime('9999-12-31T00:00:00.000Z', TRELLO_DATETIME_FORMAT)

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
    start_date = datetime(start_date.year, start_date.month, start_date.day, 0, 0, 0, 0, timezone.utc)
    end_date = card.get_due()
    end_date = datetime(end_date.year, end_date.month, end_date.day, 0, 0, 0, 0, timezone.utc)
    if start_date == INVALID_DATE_TIME or end_date == INVALID_DATE_TIME:
        return False

    date = datetime(date.year, date.month, date.day, 0, 0, 0, 0, timezone.utc)
    return start_date <= date and end_date >= date

def is_task_actual_date_in_date(card, date):
    start_date = card.get_date()
    start_date = datetime(start_date.year, start_date.month, start_date.day, 0, 0, 0, 0, timezone.utc)
    end_date = card.get_date_last_activity()
    end_date = datetime(end_date.year, end_date.month, end_date.day, 0, 0, 0, 0, timezone.utc)
    if start_date == INVALID_DATE_TIME or end_date == INVALID_DATE_TIME:
        return False

    date = datetime(date.year, date.month, date.day, 0, 0, 0, 0, timezone.utc)
    return start_date <= date and end_date >= date

def str_to_trello_format_datetime(date_str):
    if date_str == None: return INVALID_DATE_TIME 
    return datetime.strptime(date_str, TRELLO_DATETIME_FORMAT)

def calc_progress(check_list):
    if len(check_list) < 1:
        return 

    return str(int(100 * (len(list(filter(lambda x: x['state'] == 'complete', check_list))) / len(check_list)))) + '%'

def datetime_to_date(date_time):
    return datetime.date(date_time)

def next_week(date):
    current_month = calendar.Calendar().monthdatescalendar(date.year, date.month)
    day_index = -1
    week_index = 0
    for dates in current_month:
        try:
            day_index = dates.index(datetime(date.year, date.month, date.day).date())
            week_index += 1
        except ValueError:
            week_index += 1
            continue
        else:
            break

    next_week_day = INVALID_DATE_TIME.date()
    try:
        next_week_day = current_month[week_index][day_index]
    except IndexError:
        try:
            current_month = calendar.Calendar().monthdatescalendar(date.year, date.month + 1)
        except ValueError:
            current_month = calendar.Calendar().monthdatescalendar(date.year + 1, calendar.January)
        next_week_day = current_month[0][day_index]

    return datetime(next_week_day.year, next_week_day.month, next_week_day.day, 0, 0, 0, 0, timezone.utc).strftime(TRELLO_DATETIME_FORMAT)

class TrelloAPI:
    def __init__(self, json_content):
        self.__json_content = json_content
        self.__board = trello_board.Board()
        self.__check_lists = []

    def map_to_board(self):
        self.__board.set_name(self.__json_content['name'])
        self.__board.set_board_id(self.__json_content['id'])
        self.__board.set_desription(self.__json_content['desc'])
        self.__board.set_id_organization(self.__json_content['idOrganization'])
        self.__board.ser_id_member_creator(self.__json_content['idMemberCreator'])
        self.__map_to_check_lists()
        self.__map_to_cards()

    def __map_to_cards(self):
        listname = {}
        for v in self.__json_content['lists']:
            listname[v['id']] = v['name']

        members = {}
        for v in self.__json_content['members']:
            members[v['id']] = v['fullName']
        self.__board.set_members(members)

        cards = []
        for card_json in self.__json_content['cards']:
            card = trello_card.Card()
            card.set_card_id(card_json['id'])
            card.set_name(card_json['name'])
            card.set_desc(card_json['desc'])
            card.set_listname(listname[card_json['idList']])
            card.set_closed(card_json['closed'])
            card.set_date_last_activity(card_json['dateLastActivity'])

            for action in self.__json_content['actions']:
                try:
                    card_id = action['data']['card']['id']
                except KeyError:
                    continue
                if card_id == card_json['id'] :
                    card.set_date(action['date'])


            if card_json['due'] != None:
                card.set_due(card_json['due'])
            else:
                card.set_due(next_week(card.get_date()))
            card.set_due_complete(card_json['dueComplete'])
            card.set_check_list(list(filter(lambda x: x.get_id_card() == card.get_card_id(), self.__check_lists)))

            membernames = []
            for member_json in card_json['idMembers']:
                membernames.append(self.__board.get_members()[member_json])
            card.set_membernames(membernames)
            cards.append(card)

        self.__board.set_cards(cards)

    def __map_to_check_lists(self):
        for check_list in self.__json_content['checklists']:
            trello_check_list = trello_check_lists.Check_Lists()
            trello_check_list.set_check_lists_id(check_list['id'])
            trello_check_list.set_name(check_list['name'])
            trello_check_list.set_id_card(check_list['idCard'])
            trello_check_list.set_check_items(check_list['checkItems'])
            self.__check_lists.append(trello_check_list)
    
    def cards(self):
        return self.__board.get_cards()

    def board(self):
        return self.__board

    def members(self):
        return self.__board.get_members()

    def sort_cards_by_date(self):
        cards = self.__board.get_cards()
        cards.sort(key=lambda v : v.get_date())
        self.__board.set_cards(cards)

