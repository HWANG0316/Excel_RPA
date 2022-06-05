import os
from openpyxl import Workbook


def isNan(num):
    return num == num

def take_school_info(URL,count):
    import pandas as pd    
    import datetime
    import win32com.client
    from openpyxl import load_workbook

    excel = win32com.client.Dispatch("Excel.Application")   # 엑셀 색칠하기 위한 객체 생성
    excel.Visible = True

    wb = excel.Workbooks.Open(URL + '/'+ excel_file_list[count])
    ws = wb.ActiveSheet

    df = pd.read_excel(excel_file_list[count])
    
    if isNan(df.values[0][0]):
        file_name.append(df.values[0][0])
        cut_number = df.values[0][0].find('_')

        school_information[count].append(df.values[0][0][:cut_number])
        school_information[count].append(df.values[0][0][cut_number+1:])

        if(len(df.values) == 1):
            school_information[count].append(0)
        else:
            school_information[count].append(len(df.values)-3)
    else:
        file_name.append(df.values[0][1])
        cut_number = df.values[0][1].find('_')

        school_information[count].append(df.values[0][1][:cut_number])
        school_information[count].append(df.values[0][1][cut_number+1:])

        if(len(df.values) == 1):
            school_information[count].append(0)
        else:
            school_information[count].append(len(df.values)-3)
    
    
    wb.Save()
    excel.Quit()      #엑셀 닫기



print("Program Start !")
URL = os.getcwd()
path = "./"
file_list = os.listdir(path)
excel_file_list = []
file_name = []
file_name_list = []

for i in range(len(file_list)):
    if 'xlsx' in file_list[i] or 'xls' in file_list[i]:
        excel_file_list.append(file_list[i])

school_information = [[] for i in range(len(excel_file_list))]

for i in range(len(excel_file_list)):
    take_school_info(URL,i)



write_wb = Workbook()
write_ws = write_wb.active
write_ws.append(["학교명","코드명","오류건수"])

for i in range(len(school_information)):
    write_ws.append(school_information[i])


write_wb.save(URL+'/Summary.xlsx')
print("Complete")

same_name = [1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1]
same_name_list = []
same_name_count = -1

for i in range(len(file_name)):

    new_name = file_name[i] + '.xls'


    if new_name in file_name_list and new_name in same_name_list:
        os.rename(excel_file_list[i], file_name[i] + '_'+ str(same_name[same_name_count]) +'.xls')
        same_name_list.append(new_name)
        same_name[same_name_count] = same_name[same_name_count] + 1
        file_name_list.append(file_name[i] + '_'+ str(same_name[same_name_count]) +'.xls')
    elif new_name in file_name_list:
        same_name_count = same_name_count + 1
        os.rename(excel_file_list[i], file_name[i] + '_'+ str(same_name[same_name_count]) +'.xls')
        same_name_list.append(new_name)
        same_name[same_name_count] = same_name[same_name_count] + 1
        file_name_list.append(file_name[i] + '_'+ str(same_name[same_name_count]) +'.xls')
    elif new_name in excel_file_list and new_name != excel_file_list[i]:
        same_name_count = same_name_count + 1
        os.rename(excel_file_list[i], file_name[i] + '_'+ str(same_name[same_name_count]) +'.xls')
        same_name_list.append(new_name)
        same_name[same_name_count] = same_name[same_name_count] + 1
        file_name_list.append(file_name[i] + '_'+ str(same_name[same_name_count]) +'.xls')
    else:
        os.rename(excel_file_list[i], new_name)
        file_name_list.append(new_name)