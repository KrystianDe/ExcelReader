
class Account:

    def __init__(self, id_acccount, vendor, date,  sick_leaves,
                    other, total, man_days, work_from_poland, work_from_abroad):
        self.id_acccount = id_acccount
        self.vendor = vendor
        self.date = date
        self.sick_leaves = sick_leaves
        self.other = other
        self.man_days = man_days
        self.work_from_poland = work_from_poland
        self.work_from_abroad = work_from_abroad
        self.total = total
        self.projects = []

    def add_project(self, project):
        self.projects.append(project)

    def get_id_account(self):
        return self.id_acccount

    def get_date(self):
        return self.date

    def get_vendor(self):
        return self.vendor

    def get_projects(self):
        return self.projects

    def get_total(self):
        return self.total

    def get_other(self):
        return self.other

    def get_sick_leaves(self):
        return self.sick_leaves
