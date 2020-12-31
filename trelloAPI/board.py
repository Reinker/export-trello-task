class Board:
    __name = ''
    __board_id = ''
    __description = ''
    __members = {}
    __id_organization = ''
    __id_member_creator = ''
    __cards = []

    def get_name(self):
        return self.__name

    def get_cards(self):
        return self.__cards

    def get_board_id(self):
        return self.__board_id

    def get_description(self):
        return self.__description

    def get_id_organization(self):
        return self.__id_organization

    def get_id_member_creator(self):
        return self.__id_member_creator

    def get_members(self):
        return self.__members

    def set_name(self, name):
        self.__name = name

    def set_cards(self, cards):
        self.__cards = cards

    def set_board_id(self, board_id):
        self.__board_id = board_id

    def set_desription(self, description):
        self.__description = description

    def set_members(self, members):
        self.__members = members

    def set_id_organization(self, id_organization):
        self.__id_organization = id_organization
    
    def ser_id_member_creator(self, id_member_creator):
        self.__id_member_creator = id_member_creator
