class Member:
    __member_id = ''
    __full_name = ''

    def get_member_id(self):
        return self.__member_id

    def get_full_name(self):
        return self.__full_name

    def set_member_id(self, member_id):
        self.__member_id = member_id
    
    def set_full_name(self, full_name):
        self.__full_name = full_name
