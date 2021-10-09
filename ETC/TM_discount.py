# -*- coding: utf-8 -*-

import pyexcel as pe

from tkinter import *
from tkinter import filedialog

# Open Excel File

root = Tk()
root.withdraw()
root.filename = filedialog.askopenfilename(initialdir="D:/Dropbox/A-TEB/작업폴더/00.등록양식저장/", title="Open Data files", filetypes=(("data files", "*.xls;*.xlsx"), ("all files", "*.*")))


# Setting
Book = pe.get_book(file_name=root.filename)
TM = Book.sheet_by_index(0)
C = int()


# 할인율 입력

DC = int(input("Discount Rate: "))


# 할인율 적용

for i in range(2, TM.number_of_rows()):

    DC_rate = (100 - DC) / 100
    OriginalValue = int(TM.cell_value(i, 5))
    # TM.cell_value(i, 5, int(math.ceil(OriginalValue * DC_rate/100)*100))
    TM.cell_value(i, 5, int(round(OriginalValue * DC_rate, -2)))
    C += 1

    print("item #{} price {}% discounted : {} -> {}".format(C, DC, OriginalValue, TM.cell_value(i, 5)))

# TM.save_as(root.filename.split('.')[0] + '.' + root.filename.split('.')[1] + "-DCed" + root.filename.split('.')[2])
# TM.save_as(root.filename.split('.')[0] + '.' + root.filename.split('.')[1] + "-DCed.xls")

Book.save_as(root.filename.split('.')[0] + '.' + root.filename.split('.')[1] + "-DCed.xlsx")
# Book.save_as(root.filename)

print("{} items discounted".format(C))