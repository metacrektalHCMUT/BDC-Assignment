import xlrd
import openpyxl

workbook = xlrd.open_workbook("data.xlsx","rb")
sheets = workbook.sheet_names()
cmt = []
user = []
for sheet_name in sheets:
    sh = workbook.sheet_by_name(sheet_name)
    for rownum in range(sh.nrows):
        row_valaues = sh.row_values(rownum)
        user.append(row_valaues[0])
        cmt.append(row_valaues[1])

my_set =  set(open('lib.txt').read().split())

def check_Eng (text):
  lib_set = set(open('lib.txt'))
  text_lst = text.split(" ")
  lst_length = len(text_lst)
  sum = 0
  for i in text_lst:
        if (i in my_set): 
          sum = sum + 1
  if(sum/lst_length >= 0.5):
    return True
  else: 
      print(sum/lst_length)
      return False

cmt_lst = []
user_lst = []


for i in range(0, len(cmt)+1):
    if(check_Eng(cmt[i])):
      cmt_lst.append(cmt[i])
      user_lst.append(user[i])

def output_Excel(input_detail,output_excel_path):
  #Xác định số hàng và cột lớn nhất trong file excel cần tạo
  row = len(cmt_lst)
  column = 2

  #Tạo một workbook mới và active nó
  wb = openpyxl.Workbook()
  ws = wb.active
  
  #Dùng vòng lặp for để ghi nội dung từ input_detail vào file Excel
  for i in range(0,column):
    for j in range(0,row):
      v=input_detail[i][j]
      ws.cell(row=j+1, column=i+1, value=v)

  #Lưu lại file Excel
  wb.save(output_excel_path)
input_detail =[user_lst, cmt_lst]
output_excel_path= './data1.xlsx'
output_Excel(input_detail,output_excel_path)



