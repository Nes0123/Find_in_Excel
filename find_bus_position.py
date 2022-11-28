
import os
from tkinter import *
import tkinter.ttk as ttk
from tkinter import filedialog
import tkinter.messagebox as msgbox

import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# 엑셀 파일별 형식이 너무 다양해서 일단 중단

current_dir = os.getcwd()

root = Tk()
root.title("엑셀 데이터 검색 프로그램")
root.resizable(False,False)

def add_file():
    # 나중에 버스 엑셀이 xlsx인지 xls인지 보고 바꿀 것
    files = filedialog.askopenfilenames(title='엑셀 파일을 선택해주세요',
    filetypes=(("엑셀 파일 97-2003","*.xls"),("엑셀 파일","*.xlsx"),
    ("모든 파일","*.*")), initialdir="C:/")

    for file in files:
        list_file.insert(END,file)

def del_file():
    for index in reversed(list_file.curselection()):
        list_file.delete(index)


def find_txt():
    pass

# 파일 프레임
file_frame = Frame(root)
file_frame.pack(fill="x", padx=5, pady=5)

# 파일 추가 버튼 
btn_add_file = Button(file_frame, text="파일추가",
command=add_file, padx=5, pady=5, width=12)
btn_add_file.pack(side="left", padx=5, pady=5)

# 파일 삭제 버튼
btn_del_file = Button(file_frame, text="파일삭제",
command=del_file, padx=5, pady=5, width=12)
btn_del_file.pack(side="right", padx=5, pady=5)

# 리스트 프레임
list_frame = Frame(root)
list_frame.pack(expand=True, fill="both", padx=5, pady=5)

# y축 스크롤바
yscrollbar = Scrollbar(list_frame)
yscrollbar.pack(side="right", fill="y")

# x축 스크롤바
xscrollbar = Scrollbar(list_frame, orient=HORIZONTAL)
xscrollbar.pack(side="bottom", fill="x")

# 파일 리스트
list_file = Listbox(list_frame, selectmode="extended",
height=10, yscrollcommand=yscrollbar.set,
xscrollcommand=xscrollbar.set, width=60)
list_file.pack(side="left", fill="both", expand=True)

yscrollbar.config(command=list_file.yview)
xscrollbar.config(command=list_file.xview)

# 기능 프레임
func_frame = Frame(root)
func_frame.pack(fill="x", padx=5, pady=5, ipady=4)

# 검색 앤트리
entry_find = Entry(func_frame, width=20)
entry_find.pack(side="left", padx=5, pady=5, ipady=4)

# 검색 버튼
btn_find = Button(func_frame, text="검색",
command=find_txt, padx=5, pady=5)
btn_find.pack(side="left", padx=5, pady=5)

# 종료 버튼
btn_close = Button(func_frame, text="닫기",
command=root.quit, padx=5, pady=5)
btn_close.pack(side="right", padx=5, pady=5)















root.mainloop()
