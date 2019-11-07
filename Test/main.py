import Tets as Te
from tkinter import *
from openpyxl import load_workbook

if __name__ == "__main__":

    wb_val = load_workbook(filename="Kursach.xlsx",data_only=True)
    sheet_val = wb_val['Варианты']
    root = Tk()
    e = Entry(root,width = 20)
    b = Button(root, text = 'Цена за ед.пр')
    l = Label(root, bg = 'black', fg = 'white', width = 20)
    def strToSortList(event):
        s = e.get()
        s = int(s)
        a = ['B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R',
             'S','T','U','V','W','X','Y','Z']
        if s>=1:
            Num =a[s-1]
            l['text'] = sheet_val[Num+'4'].value
        else:
            Num = 'Error'
            l['text'] = 'Error Value'


    b.bind('<Button-1>', strToSortList)
    e.pack()
    b.pack()
    l.pack()
    root.mainloop()

