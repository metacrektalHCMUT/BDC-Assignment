import csv
import io
from selenium import webdriver
from selenium.common import exceptions
from selenium.webdriver.common.keys import Keys
import sys
import time
import openpyxl


browser = webdriver.Chrome(executable_path="C:/Users/HP/AppData/Roaming/Microsoft/Windows/Start Menu/Programs/Python 3.9/crawl data/chromedriver.exe")
browser.get("https://www.youtube.com/watch?v=WZMOmT3QkaE&t=4120s")
browser.maximize_window()
time.sleep(5)

title = browser.find_element_by_xpath('//*[@id="container"]/h1/yt-formatted-string').text
comment_section = browser.find_element_by_xpath('//*[@id="comments"]')

browser.execute_script("arguments[0].scrollIntoView();", comment_section)
time.sleep(7)

last_height = browser.execute_script("return document.documentElement.scrollHeight")

while True:
    browser.execute_script("window.scrollTo(0, document.documentElement.scrollHeight);")

    time.sleep(2)

    new_height = browser.execute_script("return document.documentElement.scrollHeight")
    if new_height == last_height:
        break
    last_height = new_height

browser.execute_script("window.scrollTo(0, document.documentElement.scrollHeight);")

username_elems = browser.find_elements_by_xpath('//*[@id="author-text"]')
comment_elems = browser.find_elements_by_xpath('//*[@id="content-text"]')

print("> Completed: " + title + "\n")



cmt_lst = []
user_lst = []

my_set =  set(open('lib.txt').read().split())


with io.open('result.csv', 'w', newline='', encoding="utf-16") as file:
    for username, comment in zip(username_elems, comment_elems):
          cmt_lst.append(comment.text)
          user_lst.append(username.text)
        
def output_Excel(input_detail,output_excel_path):
  #Xác định số hàng và cột lớn nhất trong file excel cần tạo
  row = len(comment_elems)
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
output_excel_path= './dstt.xlsx'
output_Excel(input_detail,output_excel_path)

browser.close()


