import Tets as Te
from tkinter import *
from openpyxl import load_workbook

if __name__ == "__main__":
    def Set_all_value(x):
        Te.Cost_per_unit = sheet_val[x+'4'].value
        Te.Cost_price = sheet_val[x+'5'].value
        Te.Raw_and_materials = sheet_val[x+'6'].value # Сырье и материалы
        Te.Wage = sheet_val[x+'7'].value # Заработная плата
        Te.Others = sheet_val[x+'8'].value # Другое
        Te.Selling_expenses = sheet_val[x+'9'].value # Коммерческие расходы
        Te.Management_expenses = sheet_val[x+'10'].value # Управленческие расходы
        Te.Stocks_r_m = sheet_val[x+'11'].value # Запасы сырья и материалов
        Te.Unfinished_production = sheet_val[x+'12'].value # Незавершенное производство
        Te.Stocks_fin_prod = sheet_val[x+'13'].value # Запасы готовой продукции
        Te.Accounts_receivable = sheet_val[x+'14'].value # Дебеторская задолженность
        Te.Account_payable = sheet_val[x+'15'].value # Кредиторская задолженность


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
        if s>=1:        #Заполнение переменных, соответсвтующими значениями Excel'a
            Num =a[s-1]
            Set_all_value(Num)

            print(Te.Cost_price)
            l['text'] = Te.Cost_per_unit
        else:
            Num = 'Error'
            l['text'] = 'Error Value'


    b.bind('<Button-1>', strToSortList)
    e.pack()
    b.pack()
    l.pack()
    root.mainloop()

