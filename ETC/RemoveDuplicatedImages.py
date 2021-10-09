# -*- coding: utf-8 -*-

import pyexcel

from tkinter import *
from tkinter import filedialog

from collections import OrderedDict

import shutil

START_ROW = 2
MAIN_IMAGE_COLUMN = 13
initial_directory = "D:/Dropbox/A-TEB/작업폴더/00.등록양식저장/"


def open_ExcelFile(title):
    """
    엑셀파일 열기
    :return: 파일명 리턴
    """
    root = Tk()
    root.withdraw()
    root.filename = filedialog.askopenfilename(initialdir=initial_directory, title=title,
                                               filetypes=(("data files", "*.csv;*.xls;*.xlsx"), ("all files", "*.*")))
    print('Data Read From << ', root.filename)
    return root.filename


def copy_File(file_name):
    """:cvar
    파일 복사
    """
    new_file_name = file_name[:-5] + "-4AT.xlsx"
    shutil.copy(file_name, new_file_name)
    print(">> File Copied : {file_name} => {new_file_name}")
    return new_file_name


def process_CopiedFile(new_file_name):

    book = pyexcel.get_book(file_name=new_file_name)
    sheet = book[0]

    for row in range(START_ROW, sheet.number_of_rows()):

        print(f"{row} : 대표이미지 {sheet[row, MAIN_IMAGE_COLUMN]}")

        image_list = []

        for column in range(MAIN_IMAGE_COLUMN+1 , MAIN_IMAGE_COLUMN+1 +5):
            if sheet[row, column] == "" or sheet[row, column] == None or sheet[row, column] == sheet[row, MAIN_IMAGE_COLUMN]:
                continue
            else:
                image_list.append(sheet[row, column])

        if len(image_list) == 0:
            image_list.append(sheet[row, MAIN_IMAGE_COLUMN])

        print(image_list)

        # image_list_set = set(image_list)
        list(OrderedDict.fromkeys(image_list))

        print(f">> {image_list}")

        i = 0
        for column in range(MAIN_IMAGE_COLUMN+1 , MAIN_IMAGE_COLUMN+1 +5):
            if i < len(image_list):
                sheet[row, column] = image_list[i]
                i += 1
                print(f"{row} {column} : {sheet[row, column]}")
            else:
                sheet[row, column] = ""
                print(f"{row} {column} : {sheet[row, column]}")

    book.save_as(new_file_name)


def main():

    file_name = open_ExcelFile('작업대상 파일 선택')
    new_file_name = copy_File(file_name)
    process_CopiedFile(new_file_name)


if __name__ == '__main__':
    # main
    main()