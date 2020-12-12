import model.card
import model.board

#プロジェクトのタスク開始日を取得する
def get_project_start_date(all_board_and_card):
    if len(all_board_and_card) < 1:
        return

    all_cards = []
    for key in all_board_and_card:
        all_cards.append(all_board_and_card[key]) 

    all_cards.sort(key=lambda v : v.get_date())
    return all_cards[0].get_date()

class TrelloAPI:
    __cards = []

    def __init__(self, content):
        self.__board = model.board.Board()
        self.__board.set_name(content['name'])

        listname = {}
        for v in content['lists']:
            listname[v['id']] = v['name']

        members = {}
        for v in content['members']:
            members[v['id']] = v['fullName']

        for card_json in content['cards']:
            card = model.card.Card()
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

