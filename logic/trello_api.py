import model.card
import model.board

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

        for v in content['cards']:
            card = model.card.Card()
            card.set_card_id(v['id'])
            card.set_name(v['name'])
            card.set_desc(v['desc'])
            card.set_listname(listname[v['idList']])
            card.set_closed = v['closed']
            card.set_date_last_activity = v['dateLastActivity']
            card.set_due = v['due']
            card.set_due_complete = v['dueComplete']

            membernames = []
            for m in v['idMembers']:
                membernames.append(members[m])
            card.set_membernames(membernames)

            self.__cards.append(card)
        
        self.__board.set_cards(self.__cards)

    def get_cards(self):
        return self.__cards

    def get_board(self):
        return self.__board
