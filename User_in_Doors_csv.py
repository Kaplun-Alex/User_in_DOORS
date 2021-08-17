from tkinter import *
from tkinter import ttk
import tkinter.font
import pypyodbc
import csv

win = tkinter.Tk()
win.title("Create number")
win.geometry('600x450+100+100')
ButtonFont = tkinter.font.Font(family='Hervetica', size=10, weight='bold')
BigFont = tkinter.font.Font(family='Hervetica', size=12, weight='bold')


def chek_doors_list():
    try:
        cheksqlquery = (
                """select colPointName FROM dbo.tblControlPoints""")  # - Проверка по индексу таблицы
        print(cheksqlquery)
        cursor.execute(cheksqlquery)
        results = cursor.fetchall()
        ls_of_doors = []
        for i in results:
            ls_of_doors.append(i[0])
        print(results)
        print(ls_of_doors)
        select_dor.config(value=ls_of_doors)
    except NameError:
        stEntry.delete(0, END)
        stEntry.insert(0, "ДО БАЗИ ПІД'ЄДНАЙСЯ ТЕЛЕПНЮ!")


def chbasestb():
    closeconection()
    global cursor
    global connection
    db_host = '172.25.7.33'
    db_name = 'NextAccessAll'
    db_user = 'sa'
    db_password = 'Admin25112012'
    STBLabel.config(bg="#00cc00")
    try:
        connection = pypyodbc.connect(
            'Driver={SQL Server};''Server=' + db_host + ';''Database=' + db_name + ';''uid=' + db_user + ';''pwd=' + db_password + ';')
        cursor = connection.cursor()
        stEntry.delete(0, END)
        stEntry.insert(0, "Зв'язок з STB встановлено")
    except:
        stEntry.delete(0, END)
        stEntry.insert(0, "Халепа! Зв'язок з STB відсутній")
        STBLabel.config(bg="#FF0000")
    chek_doors_list()
    return connection, cursor


def closeconection():
    try:
        connection.close()
    except NameError:
        print("Зв'язку не було")


def exitProgram():
    try:
        connection.close()
        win.destroy()
    except NameError:
        win.destroy()


def find_user_in_door():
    try:
        door = select_dor.get()
        mod_dor = "'"+door+"'"
        print(door, mod_dor, type(door))
        cheksqlquery = (
                """SELECT     TOP (100) PERCENT 
                            dbo.vwAllCardholders.colAccountNumber,
                            dbo.vwAllCardholders.colSurname, 
                            dbo.vwAllCardholders.EmpName, 
                            dbo.vwAllCardholders.colStatus, 
                            dbo.tblControlPoints.colPointName
                FROM dbo.tblWorkCycles INNER JOIN
                            dbo.tblCycleDays ON dbo.tblWorkCycles.colCycleID = dbo.tblCycleDays.colCycleID INNER JOIN
                            dbo.tblDayPeriods ON dbo.tblCycleDays.colDayID = dbo.tblDayPeriods.colDayID INNER JOIN
                            dbo.tblPointsAccRights ON dbo.tblDayPeriods.colPeriodID = dbo.tblPointsAccRights.colPeriodID INNER JOIN
                            dbo.tblControlPoints ON dbo.tblPointsAccRights.colPointNumber = dbo.tblControlPoints.colPointNumber INNER JOIN
                            dbo.vwAllCardholders ON dbo.tblWorkCycles.colCycleID = dbo.vwAllCardholders.colCycleID
                WHERE (dbo.vwAllCardholders.CardStatus=1) and 
                            (dbo.tblPointsAccRights.colRights = 1) and 
                            (dbo.tblControlPoints.colPointName="""+mod_dor+""")
                            ORDER BY dbo.tblControlPoints.colPointName""")  # - Проверка по индексу таблицы
        print(cheksqlquery)
        cursor.execute(cheksqlquery)
        results = cursor.fetchall()
        print(results)
        with open('Спис_сотр_прописанных_в_дверь'+door+'.csv', 'w', newline='') as file:  # ну це ахуэнно, ГЕ?
            writer = csv.writer(file, delimiter=';', dialect='excel')
            writer.writerows(results)
            file.close()
    except NameError:
        stEntry.delete(0, END)
        stEntry.insert(0, "Сталася невідома хрінь! З'эднався з базою або обери точку доступу")


stEntry = Entry(win, font=ButtonFont, bg='#C0C0C0')
stEntry.pack()
stEntry.place(x=20, y=420, height=20, width=460)
stEntry.delete(0, END)
stEntry.insert(0, 'Все заєбісь!')

select_dor = ttk.Combobox(win, values=['Не підглядуй!'], height=20, width=60, font=BigFont)
select_dor.pack()
select_dor.place(x=20, y=150)

exitButton = Button(win, text="ВИХІД", font=ButtonFont, command=exitProgram, height=2, width=8)
exitButton.pack()
exitButton.place(x=515, y=400)

chBratButton = Button(win, text='Перевірити', font=ButtonFont, command=find_user_in_door, height=1, width=20)
chBratButton.pack()
chBratButton.place(x=200, y=300)

numberLabel = Label(win, text='Обери необхшдні двері братуха!', font=ButtonFont, justify=LEFT, width=30)
numberLabel.pack()
numberLabel.place(x=10, y=125)

STBLabel = Label(win, text='БАЗА STB', font=ButtonFont, width=12)
STBLabel.pack()
STBLabel.place(x=235, y=20)

statusLabel = Label(win, text='Статус', font=ButtonFont, width=12)
statusLabel.pack()
statusLabel.place(x=0, y=390)

chSTBBase = Button(win, text="З'ЄДНАТИ", font=ButtonFont, command=chbasestb, height=1, width=10)
chSTBBase.pack()
chSTBBase.place(x=240, y=50)

mainloop()
