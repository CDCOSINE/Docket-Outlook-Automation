# -*- coding: utf-8 -*-
"""
Created on Wed Apr 22 12:02:58 2020

@author: Lumenci 3
"""


import win32com.client
#from colorama import Fore
from termcolor import colored
import dateutil.parser as dparser
import openpyxl
import datetime

import os
desktop = os.path.join(os.path.join(os.path.expanduser('~')), 'Desktop')
print(desktop)



import PySimpleGUI as sg




sg.theme('DarkAmber')   
layout1 = [[sg.Text('Enter Email/Name to look'),sg.InputText()],
            [sg.Text('Enter keywords seprated by Comma'),sg.InputText()],
            [sg.Button('Ok')] ]

window = sg.Window('Docket-Outlook-Automation',layout1)
a,b = window.read()
emaill = str(b[0])
keywordslist = str(b[1])
print('Email',emaill)
print('Keyword',keywordslist)
window.close()

layout = [  [sg.Text('Docket-Outlook-Automation')],
            [sg.Text('Enter Start Year'), sg.InputText()],
            [sg.Text('Enter End Year'), sg.InputText()],
            [sg.Text('Enter Start Month'), sg.InputText()],
            [sg.Text('Enter End Month'), sg.InputText()],
            [sg.Text('Enter Start Date'), sg.InputText()],
            [sg.Text('Enter Start Year'), sg.InputText()],
            [sg.Button('Ok'), sg.Button('Cancel')] ]

window = sg.Window('Docket-Outlook-Automation', layout)
a,b = window.read()
sty = int(b[0])
eny = int(b[1])
stm = int(b[2])
enm = int(b[3])
std = int(b[4])
edt = int(b[5])
start = datetime.date(sty,stm,std)
end = datetime.date(eny,enm,edt)
print("startt ",start)
print("end",end)
window.close()




outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
add = emaill
stringk = keywordslist
#
#dtstart = int(input("Enter start date"))
#dtend = int(input("Enter end date"))
#datestamp1 = datetime.date(startyear,startmonth,dtstart)
#datestamp2 = datetime.date(endyear,endmonth,dtend)
key = stringk.split(",")
inbox = outlook.GetDefaultFolder(6)

dt = str(datetime.date.today())
wb = openpyxl.Workbook()
ws =  wb.active
ws.title = "Sheet1"
wb["Sheet1"]['A'+str(1)] = "Email Scrapped: " + add
wb["Sheet1"]['B'+str(1)] = "Key Searched " + stringk
wb["Sheet1"]['A'+str(2)] = "Keyy matched"
wb["Sheet1"]['B'+str(2)] = 'Campaign'
wb["Sheet1"]['C'+str(2)] = 'Docket Update'
wb["Sheet1"]['D'+str(2)] = 'Campaign Link'
wb["Sheet1"]['E'+str(2)] = 'Docket Link'
wb["Sheet1"]['F'+str(2)] = 'Docket Date extracted from Docket Text'
wb.save(desktop+"/" +dt+stringk+add+ ".xlsx")
messages = inbox.Items
found = ""
length=len(messages)
work = openpyxl.load_workbook(desktop+"/" +dt+stringk+add+ ".xlsx")
print("Yes")
sheet = work["Sheet1"]
k = 3
for i in range(length-1,-1,-1):
    print(datetime.date.today())
    if messages[i].Class == 43 and messages[i].senton.date()>=start and messages[i].senton.date()<=end:
        try:
            a=messages[i].Sender.GetExchangeUser().PrimarySmtpAddress
        except:     
            a=messages[i].SenderEmailAddress
        if add.lower() in a.lower():
                print("add, started---------------------------------")
                liness = messages[i].body.lower().split("\n")
                lastcamp = ""
                lastlink = ""
                for c,lin in enumerate(liness):
                    if "campaign:" in str(lin).lower():
                        print("")
                        lis = str(lin).split("<")
                        lastcamp =  lis[0].split(": ")[1]
                    for keyyy in key:
                        if keyyy in lin:
                            found = str(lin)
                            sheet['A'+str(k)] = keyyy
                            print("######################")
                            print("Campaign: ",lastcamp)
                            print("keyword: ",found)
                            print("Link ",lis[1].split(">")[0])
                            sheet['B'+str(k)] = lastcamp
                            sheet['C'+str(k)] = found
                            try:
                                sheet['D'+str(k)] = lis[1].split(">")[0]
                            except:
                                print("..")
                            try:
                                l1 = found.split("<")[1].split(">")[0]
                                sheet['E'+str(k)] = l1
                            except:
                                print('00')
                            try:
                                dat = str(dparser.parse(found,fuzzy=True))
                                sheet['F'+str(k)] = dat
                            except:
                                print('00')
                            k = k+1
                            print("######################")
                        
                    print("Over---------------------------------")
print(desktop+"/" +dt+stringk+add+ ".xlsx")
work.save(desktop+"/" +dt+stringk+add+ ".xlsx")
