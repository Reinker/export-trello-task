class Check_Lists:
    __check_lists_id = ''
    __name = ''
    __id_card = ''
    __check_items = []
    
    def set_check_lists_id(self, check_lists_id):
        self.__check_lists_id = check_lists_id

    def set_name(self, name):
        self.__name = name

    def set_id_card(self, id_card):
        self.__id_card = id_card

    def set_check_items(self, check_items):
        self.__check_items = check_items

    def get_check_lists_id(self):
        return self.__check_lists_id
    
    def get_name(self):
        return self.__name
    
    def get_id_card(self):
        return self.__id_card

    def get_check_items(self):
        return self.__check_items
