class Board:
    __name = ''
    __board_id = ''
    __description = ''
    __cards = []

    def get_name(self):
        return self.__name

    def get_cards(self):
        return self.__cards

    def get_board_id(self):
        return self.__board_id

    def get_description(self):
        return self.__description

    def set_name(self, name):
        self.__name = name

    def set_cards(self, cards):
        self.__cards = cards

    def set_board_id(self, board_id):
        self.__board_id = board_id

    def set_desription(self, description):
        self.__description = description
