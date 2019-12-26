from tkinter import *
from ExcelReader import Excel_Reader
import os



month_values = [x for x in range(1, 13)]

fileDir = os.path.dirname(os.path.abspath(__file__))
parentDir = os.path.dirname(fileDir)
newPath = os.path.join(parentDir, 'IS')
print(newPath)
year_values = [name for name in os.listdir(newPath) if os.path.isdir(newPath+'\\'+name)]
app = Tk()

app.title('Excel Reader')
app.geometry('250x150')

mon = None
yea = None

def callback_m(selection):
    print(selection)
    global mon
    mon = selection

def callback_y(selection):
    print(selection)
    global yea
    yea = selection

part_month = StringVar(app)
part_month.set('None')
part_month = Label(app, text='Miesiąc', font=('bold',14), pady=10)
part_month.grid(row=0, column=0)
# part_month = Entry(app, textvariable=part_month)
# part_month.grid(row=0,column=1)
part_month = StringVar(app)
part_month.set('None')
part_month = OptionMenu(app, part_month, *month_values
                            , command=callback_m)

part_month.grid(row=0,column=1)

part_year = StringVar()
part_year = Label(app, text='Rok', font=('bold',14), pady=10)
part_year.grid(row=1, column=0)
# part_year = Entry(app, textvariable=part_year)

part_year = StringVar(app)
part_year.set('None')
part_year = OptionMenu(app, part_year, *year_values
                            , command=callback_y)
part_year.grid(row=1,column=1)

def generate_file():
    month = mon
    year = yea
    Excel_Reader(month, year, newPath).reading_excels()


generate_btn = Button(app, text="Generuj plik csv", width=12, command=generate_file)
generate_btn.grid(row=3,column=1, pady=15)




app.mainloop()
