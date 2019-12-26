class Project:

    def __init__(self, name_of_project, work_from_abroad, work_from_poland,
                     total, man_days):
        self.name_of_project = name_of_project
        self.work_from_abroad = work_from_abroad
        self.work_from_poland = work_from_poland
        self.total = total
        self.man_days = man_days
        # self.services = services
        # self.sick_leaves = sick_leaves
        # self.other = other

    def get_name_of_project(self):
        return self.name_of_project

    def get_work_from_abroad(self):
        return self.work_from_abroad

    def get_work_from_poland(self):
        return self.work_from_poland

    def get_total(self):
        return self.total    

    def get_man_days(self):
        return self.man_days    

    # def get_services(self):
    #     return self.services

    # def get_other(self):
    #     return self.other

    # def get_sick_leaves(self):
    #     return self.sick_leaves


