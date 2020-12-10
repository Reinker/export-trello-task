class Card :
    __card_id = ''
    __name = ''
    __desc = ''
    __date = ''
    __listname = ''
    __closed = False
    __membernames = []
    __date_last_activity = ''
    __due = ''
    __due_complete = False
    
    def set_card_id(self, card_id):
        self.__card_id = card_id

    def set_name(self, name):
        self.__name = name

    def set_desc(self, desc):
        self.__desc = desc

    def set_listname(self, listname):
        self.__listname = listname

    def set_closed(self, closed):
        self.__closed = closed

    def set_membernames(self, membernames):
        self.__membernames = membernames

    def set_date_last_activity(self, date_last_activity):
        self.__date_last_activity = date_last_activity

    def set_due(self, due):
        self.__due = due

    def set_due_complete(self, due_complete):
        self.__due_complete = due_complete

    def set_date(self, date):
        self.__date = date

    def get_card_id(self):
        return self.__card_id

    def get_name(self):
        return self.__name

    def get_desc(self):
        return self.__desc

    def get_listname(self):
        return self.__listname

    def get_closed(self):
        return self.__closed

    def get_membernames(self):
        return self.__membernames

    def get_date_last_activity(self):
        return self.__date_last_activity

    def get_due(self):
        return self.__due

    def get_due_complete(self):
        return self.__due_complete
    
    def get_date(self):
        return self.__date
