import openpyxl
import os
import copy
import shutil
from openpyxl import load_workbook
import win32com.client
import numpy


def take_criteria_name():
    import pandas as pd    
    import datetime
    import win32com.client
    from openpyxl import load_workbook

    df = pd.read_excel(criteria_file[0])
    data_code = df['학교명'].tolist()
    
    return data_code


print("Program Start !")
URL = os.getcwd()

compare_file_list = []
criteria_file  = []



path = "./2. 공간현황_조사전_전체/"
file_list = os.listdir(path)
compare_file_list = copy.deepcopy(file_list)

path = "./"
file_list = os.listdir(path)


for i in range(len(file_list)):
    if 'xlsx' in file_list[i] or 'xls' in file_list[i]:
        criteria_file.append(file_list[i])




comp_file_value_list = []
criteria_file_value_list = []
name_value_list = []


for i in range(len(compare_file_list)):
    cut_number = compare_file_list[i].find('_')
    cut_number2 = compare_file_list[i].find('조사전')
    name_value_list.append(compare_file_list[i][cut_number + 1: cut_number2-1])

criteria_name = take_criteria_name()


move_file = []
for i in range(len(name_value_list)):
    if name_value_list[i] in criteria_name:
        move_file.append(compare_file_list[i])


if not(os.path.isdir("검증완료_학교명")):
    os.makedirs(os.path.join("검증완료_학교명")) 

src = './2. 공간현황_조사전_전체/'
dir = './검증완료_학교명/'


excel = win32com.client.Dispatch("Excel.Application")   # 엑셀 색칠하기 위한 객체 생성
excel.Visible = False

wb = excel.Workbooks.Open(URL + '/'+ criteria_file[0])
ws = wb.ActiveSheet

ws.Cells(1,4).Value = "학교명"

count = 0
for i in range(len(name_value_list)):
    if name_value_list[i] in criteria_name:
        ws.Cells(criteria_name.index(name_value_list[i]) + 2,4).Value = "O"
        count = count + 1

wb.Save()
excel.Quit() 


for i in range(len(move_file)):
    filename = move_file[i]
    shutil.move(src + filename, dir + filename)
    