import os
import openpyxl
import csv
from Account import Account
from Project import Project


class Excel_Reader:

    def __init__(self):
        self.excel_files = self.get_excel_files(os.getcwd())

    def get_excel_files(self, path):
        files = []
        for r, d, f in os.walk(path):
            for file in f:
                if '.xlsx' in file:
                    files.append(os.path.join(r, file))

        return files


    def reading_excels(self):
        with open('result.csv', 'w', newline='') as csvfile:
            fieldnames = ['Id','Month', 'Vendor', 'Main Project', 'Hours Poland',
                    'Hours Abroad', 'Total hours', 'Sick Leaves', 'Other']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

            writer.writeheader()

            for f in self.excel_files:
                wb = openpyxl.load_workbook(f)

                sheet = wb['Sheet1']

                user_id = sheet.cell(row=2, column=1).value
                vendor = sheet.cell(row=4, column=1).value
                date = sheet.cell(row=4, column=2).value
                total = sheet.cell(row=34, column=5).value
                man_days = sheet.cell(row=35, column=5).value
                work_from_poland = sheet.cell(row=36, column=5).value
                work_from_abroad = sheet.cell(row=37, column=5).value
                services = sheet.cell(row=38, column=5).value
                sick_leaves = sheet.cell(row=39, column=5).value
                other = sheet.cell(row=40, column=5).value

                account = Account(user_id, vendor, date, services, sick_leaves,
                                    other, total, man_days, work_from_poland,
                                     work_from_abroad)

                for i in range(6,10):
                    name_of_project = sheet.cell(row=3, column=i).value

                    if name_of_project == 'N/A' or name_of_project == '':
                        break

                    total_project = sheet.cell(row=34, column=i).value
                    man_days_project = sheet.cell(row=35, column=i).value
                    work_from_poland_project = sheet.cell(row=36, column=i).value
                    work_from_abroad_project = sheet.cell(row=37, column=i).value
                    services_project = sheet.cell(row=38, column=i).value
                    sick_leaves_project = sheet.cell(row=39, column=i).value
                    other_project = sheet.cell(row=40, column=i).value


                    account.add_project(Project(name_of_project,
                                                work_from_abroad_project,
                                                work_from_poland_project,
                                                total_project,
                                                services_project,
                                                sick_leaves,
                                                other_project))

                for project in account.get_projects():

                    writer.writerow({'Id': account.get_id_account(),
                                    'Month': account.get_date().date(),
                                    'Vendor': account.get_vendor(),
                                    'Main Project': project.get_name_of_project(),
                                    'Hours Poland': project.get_work_from_poland(),
                                    'Hours Abroad':project.get_work_from_abroad(),
                                    'Total hours':project.get_total(),
                                    'Sick Leaves':project.get_sick_leaves(),
                                    'Other': project.get_other()})





Excel_Reader().reading_excels()
