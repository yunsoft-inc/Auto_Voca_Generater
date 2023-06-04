#It use CamBridge Dictionary.
#https://dictionary.cambridge.org/ko/%EC%82%AC%EC%A0%84/%EC%98%81%EC%96%B4-%ED%95%9C%EA%B5%AD%EC%96%B4/

# tkinter setup
from tkinter import *
from tkinter import ttk
from time import sleep
import tkinter.messagebox
from tkinter import filedialog

#Selneium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.common.by import By
#Excel
import win32com.client

import os
import sys

# tkinter 객체 생성
window = Tk()
window.geometry('500x150+300+300')
window.title('Voca Auto Generater')
window.resizable(0,0)

def Load():
    
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Quit()
    excel.Visible = False #show progress
    filename = filedialog.askopenfilename(initialdir="/", title="Select file",
                                          filetypes=(("Excel files", "*.xlsx"),
                                          ("all files", "*.*")))
    print(filename)
    wb = excel.Workbooks.Open(r"" + filename)
    start()

def start():

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False #show progress
    datalen = int(excel.Cells(1, 2).Value)
    datalenstr = str(round(excel.Cells(1, 2).Value))
    l = 1
    tofind = [""] * datalen
    for i in range(1,datalen+1):
        data = excel.Cells(l, 1).Value
        tofind[l-1] = data
        l = l + 1
    tkinter.messagebox.showinfo("Info", datalenstr + "개의 단어가 불러와졌습니다.")
    excel.Quit()
    prog_bar["value"] = 2
    prog_bar.update()
    #Max Meaing number
    excel.Visible = True #show progress
    #init chrome driver
    op = webdriver.ChromeOptions()
    op.add_experimental_option("excludeSwitches", ["enable-logging"])
    op.add_argument('headless')
    driver = webdriver.Chrome('./chromedriver.exe',options=op)

    prog_bar["value"] = 5
    prog_bar.update()
    wb = excel.Workbooks.Add() #add workbooks
    ws = wb.Worksheets("Sheet1") #setting worksheet
    numbering = 1
    prog_bar["value"] = 10
    prog_bar.update()
    for n in tofind:
        driver.get('https://dictionary.cambridge.org/ko/%EC%82%AC%EC%A0%84/%EC%98%81%EC%96%B4-%ED%95%9C%EA%B5%AD%EC%96%B4/' + n)
        driver.refresh()
        #find meaning
        meaning_list = ''
        meanings = driver.find_elements(By.XPATH, '//*[@id="page-content"]/div[2]/div[2]/div/span/div[1]/div[3]/div/div[2]/div/div[3]/span')
        ws.cells(numbering,1).Value = n
        meannumbering = 1
        prog_bar["value"] = 13
        prog_bar.update()
        for i in meanings:
            if(meannumbering > int(max.get())):
                break
            meaning_list = meaning_list + ', ' + i.text
            meannumbering = meannumbering + 1
        prog_bar["value"] = 17
        prog_bar.update()
        ws.cells(numbering,2).Value = meaning_list[2:]
        numbering = numbering + 1          
    prog_bar["value"] = 20
    prog_bar.update()
    tkinter.messagebox.showinfo("Info", "생성완료.\n엑셀파일 저장하거나, 복사하십시오.")

max, log = StringVar(), StringVar()
max.set(2)
ttk.Label(window, text = "단어장 엑셀파일 선택 : ").grid(row = 0, column = 0, padx = 10, pady = 10)
ttk.Label(window, text = "최대 뜻 수 : ").grid(row = 1, column = 0, padx = 10, pady = 10)
ttk.Button(window, text = "불러오기", command = Load).grid(row = 0, column = 1, padx = 10, pady = 10)
ttk.Entry(window, textvariable = max).grid(row = 1, column = 1, padx = 10, pady = 10)
ttk.Progressbar(window, maximum=100, length=150)
prog_bar = ttk.Progressbar(window, length=350, maximum=20)
prog_bar.grid(row=2, column=0, padx=10, pady=10)

window.mainloop()
