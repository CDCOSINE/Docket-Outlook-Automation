# -*- coding: utf-8 -*-
"""
Created on Wed Apr 22 12:02:58 2020

@author: Lumenci 3
"""

import pip

def install(package):
    if hasattr(pip, 'main'):
        pip.main(['install', package])
    else:
        pip._internal.main(['install', package])

# Example
#install('PySimpleGUI')


import win32com.client
import dateutil.parser as dparser
import openpyxl
import datetime
import pandas as pd
import os
desktop = os.path.join(os.path.join(os.path.expanduser('~')), 'Desktop')
print(desktop)



import PySimpleGUI as sg




sg.theme('DarkAmber')   
layout1 = [[sg.Text('Enter Email/Name to look'),sg.InputText()],
            [sg.Text('On the next window Select Keywords csv File')],
            [sg.Button('Ok')] ]


window = sg.Window('Docket-Outlook-Automation',layout1)
a,b = window.read()
emaill = str(b[0])
print('Email',emaill)
window.close()


import tkinter as tk

from tkinter import filedialog

root = tk.Tk()
root.withdraw()

filr_path = filedialog.askopenfilename(title='Select Keyword CSV File to choose for Keywords',filetypes = (("CSV Files","*.csv"),))


newkeyss = pd.read_csv(filr_path)
newkeyss = list(newkeyss['Keywords'])
print('abc',newkeyss)

layout = [  [sg.Text('Choose Start Date')],
            [sg.In(key='-CAL-', enable_events=True, visible=False), sg.CalendarButton('Select Start Date', target='-CAL-', pad=None, font=('MS Sans Serif', 10, 'bold'),
                 key='_CALENDAR_', format=('%d %m, %Y'))]
           ]

window = sg.Window('Start Date', layout,size=(250,150))

event, values = window.read()
print(event,values)
window.close()

svalues = values
print(svalues['-CAL-'])
splitval = svalues['-CAL-'].split(' ')
yearrr = int(splitval[2])
monthhh = int(splitval[1].replace(',',''))
dateee = int(splitval[0])
sjk = datetime.date(yearrr,monthhh,dateee)
print(sjk)



layout = [  [sg.Text('Choose End Date')],
            [sg.In(key='-CAL-', enable_events=True, visible=False), sg.CalendarButton('Select End Date', target='-CAL-', pad=None, font=('MS Sans Serif', 10, 'bold'),
                 key='_CALENDAR_', format=('%d %m, %Y'))],
            [sg.Exit()]]

window = sg.Window('End Date', layout,size=(250,150))
event, values = window.read()
print(event,values)
window.close()




evalues = values
print(evalues['-CAL-'])
splitval = evalues['-CAL-'].split(' ')
yearrr = int(splitval[2])
monthhh = int(splitval[1].replace(',',''))
dateee = int(splitval[0])
ejk = datetime.date(yearrr,monthhh,dateee)
print(ejk)



#layout = [  [sg.Text('Docket-Outlook-Automation')],
#            [sg.Text('Enter Start Year'), sg.InputText()],
#            [sg.Text('Enter End Year'), sg.InputText()],
#            [sg.Text('Enter Start Month'), sg.InputText()],
#            [sg.Text('Enter End Month'), sg.InputText()],
#            [sg.Text('Enter Start Date'), sg.InputText()],
#            [sg.Text('Enter End Date'), sg.InputText()],
#            [sg.Button('Ok'), sg.Button('Cancel')] ]
#
#window = sg.Window('Docket-Outlook-Automation', layout)
#a,b = window.read()
#sty = int(b[0])
#eny = int(b[1])
#stm = int(b[2])
#enm = int(b[3])
#std = int(b[4])
#edt = int(b[5])
#start = datetime.date(sty,stm,std)
#end = datetime.date(eny,enm,edt)
#print("startt ",start)
#print("end",end)
#window.close()

start = sjk
end = ejk


from termcolor import colored

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
add = emaill
#stringk = keywordslist
##
##dtstart = int(input("Enter start date"))
##dtend = int(input("Enter end date"))
##datestamp1 = datetime.date(startyear,startmonth,dtstart)
##datestamp2 = datetime.date(endyear,endmonth,dtend)
#key = stringk.split(",")
stringk = ' , '.join(newkeyss)
print(stringk)

key = newkeyss
print(key)
inbox = outlook.GetDefaultFolder(6)

dt = str(datetime.date.today())
wb = openpyxl.Workbook()
ws =  wb.active
ws.title = "Sheet1"
wb["Sheet1"]['A'+str(1)] = "Email Scrapped: " + add
wb["Sheet1"]['B'+str(1)] = "Key Searched " + stringk
wb["Sheet1"]['A'+str(2)] = "Keyy matched"
wb["Sheet1"]['B'+str(2)] = 'Campaign'
wb["Sheet1"]['C'+str(2)] = 'Case Name'
wb["Sheet1"]['D'+str(2)] = 'Case Number'
wb['Sheet1']['E'+str(2)] = 'Date of Email'
wb["Sheet1"]['F'+str(2)] = 'Docket Update'
wb["Sheet1"]['G'+str(2)] = 'Campaign Link'
wb["Sheet1"]['H'+str(2)] = 'Docket Link'
wb["Sheet1"]['I'+str(2)] = 'Docket Date extracted from Docket Text'
wb.save(desktop+"/" +dt+stringk+add+ ".xlsx")
messages = inbox.Items
found = ""
length=len(messages)
work = openpyxl.load_workbook(desktop+"/" +dt+stringk+add+ ".xlsx")
print("Yes")
sheet = work["Sheet1"]
k = 3
outputwin = [
    [sg.Output(size=(78,20))]
]
progresslayout = [[sg.Text('Processing File')],
          [sg.ProgressBar(length, orientation='h', size=(20, 20), key='progressbar')],
          [sg.Frame('Output', layout = outputwin)],
          [sg.Button('Ok')],
          [sg.Cancel()]]
window = sg.Window('Processing File', progresslayout)
progress_bar = window['progressbar']

bs = "\033[1m"
be = "\033[0;0m"
for i in range(length-1,-1,-1):
    proc = length-i
    event, values = window.read(timeout=10)
    if event == 'Cancel'  or event == sg.WIN_CLOSED:
        break
  # update bar with loop value +1 so that bar eventually reaches the maximum
    progress_bar.UpdateBar(proc + 1)
#    proc = length-i
#    sg.theme('DarkAmber')   
#    layoutt1 = [[sg.Text('Processing Element{}/{}'.format(proc,length))]]
#    window = sg.Window('Docket-Outlook-Automation',layoutt1)
#    a = window.read()
#    window.close()
    if messages[i].Class == 43 and messages[i].senton.date()>=start and messages[i].senton.date()<=end:
        try:
            a=messages[i].Sender.GetExchangeUser().PrimarySmtpAddress
        except:
            a=messages[i].SenderEmailAddress
        if add.lower() in a.lower():
                liness = messages[i].body.lower().split("\n")
                lastcamp = ""
                lastlink = ""
                casenum = ''
                for c,lin in enumerate(liness):
                    if "campaign:" in str(lin).lower():
                        lis = str(lin).split("<")
                        lastcamp =  lis[0].split(": ")[1]
                        casename = str(liness[c+4]).split('<')[0]
                        print('')
                        print('New Case: ',lastcamp,' -- ',casename)
                        print('')
                    if '-cv-' in str(lin).lower():
                        casenum = str(lin).lower().split(' ')[0]
                    for keyyy in key:
                        keyyy = keyyy.lower()
                        if keyyy == 'so':
                            lin = lin.upper()
                            keyyy = 'SO'
                        if keyyy == 'po':
                            lin = lin.upper()
                            keyyy = 'PO'
                        donot = 'parties terminated'
                        if keyyy in lin.lower():
                            print('')
                            print('(Key matched = {}) == in line ==  {}'.format(keyyy,lin))
                            print('')
                            found = str(lin)
                            sheet['A'+str(k)] = keyyy
                            #print("Link ",lis[1].split(">")[0])
                            sheet['B'+str(k)] = lastcamp
                            sheet['C'+str(k)] = casename
                            sheet['D'+str(k)] = casenum
                            sheet['E'+str(k)] = messages[i].senton.date()
                            sheet['F'+str(k)] = found
                            try:
                                sheet['G'+str(k)] = lis[1].split(">")[0]
                            except:
                                print("...",end=' ')
                            try:
                                l1 = found.split("<")[1].split(">")[0]
                                sheet['H'+str(k)] = l1
                            except:
                                print('...',end=' ')
                            try:
                                dat = str(dparser.parse(found,fuzzy=True))
                                sheet['I'+str(k)] = dat
                            except:
                                print('...',end=' ')
                            k = k+1
print('OVER------------------------------------------------------------')

                      
aa= window.read()
window.close()





print(desktop+"/" +dt+stringk+add+ ".xlsx")
work.save(desktop+"/" +dt+stringk+add+ ".xlsx")
sg.theme('DarkAmber')   
layoutt = [[sg.Text('Process Over')],
            [sg.Text('Get your file at:{}'.format(desktop+"/" ))],
            [sg.Button('Ok')]]
window = sg.Window('Docket-Outlook-Automation',layoutt,size=(300,110))
a = window.read()
window.close()
