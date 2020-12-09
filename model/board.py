class Board:
    __name = ''
    __cards = []

    def get_name(self):
        return self.__name

    def get_cards(self):
        return self.__cards

    def set_name(self, name):
        self.__name = name

    def set_cards(self, cards):
        self.__cards = cards
