from abc import ABC, abstractmethod


class Services(ABC):

    @abstractmethod
    def db_create_table(self):
        pass

    @abstractmethod
    def db_add_employee(self):
        pass

    @abstractmethod
    def db_update_employee(self):
        pass

    @abstractmethod
    def db_delete_employee(self):
        pass

    @abstractmethod
    def db_get_one_employee(self):
        pass

    @abstractmethod
    def db_get_all_employee(self):
        pass

    @abstractmethod
    def xl_add_employee(self, emp):
        pass

    @abstractmethod
    def xl_update_employee(self, emp):
        pass

    @abstractmethod
    def xl_delete_employee(self, eid):
        pass

    @abstractmethod
    def xl_get_one_employee(self, eid):
        pass

    @abstractmethod
    def xl_get_all_employee(self):
        pass
