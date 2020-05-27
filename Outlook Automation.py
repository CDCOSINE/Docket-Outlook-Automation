# -*- coding: utf-8 -*-
"""
Created on Wed Apr 22 12:02:58 2020

@author: Lumenci 3
"""


import win32com.client
#from colorama import Fore
from termcolor import colored
import openpyxl
import datetime
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
add = input("Enter email address to Look for ")
key = input("Enter keyword ")
inbox = outlook.GetDefaultFolder(6)

dt = str(datetime.date.today())
wb = openpyxl.Workbook()
ws =  wb.active
ws.title = "Sheet1"
wb["Sheet1"]['A'+str(1)] = "Email Scrapped: " + add
wb["Sheet1"]['B'+str(1)] = "Key Searched " + key
wb["Sheet1"]['A'+str(2)] = 'Campaign'
wb["Sheet1"]['B'+str(2)] = 'Docket Update'
wb["Sheet1"]['C'+str(2)] = 'Campaign Link'
wb["Sheet1"]['D'+str(2)] = 'Docket Link'
wb.save('C:/Users/Lumenci 3/Desktop/'+dt+key+add+'.xlsx')
messages = inbox.Items
found = ""
length=len(messages)
work = openpyxl.load_workbook('C:/Users/Lumenci 3/Desktop/'+dt+key+add+'.xlsx')
print("Yes")
sheet = work["Sheet1"]
k = 3
for i in range(length-1,-1,-1):
    
    print(datetime.date.today())
    if messages[i].Class == 43 and messages[i].senton.date()==datetime.date.today():
        try:

            a=messages[i].Sender.GetExchangeUser().PrimarySmtpAddress
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
                        #try:
                            #lastlink = list[1]
                        #except:
                            #lastlink = "No link attached"
                    if key in lin:
                        found = str(lin)
                        print("######################")
                        print("Campaign: ",lastcamp)
                        print("keyword: ",found)
                        print("Link ",lis[1].split(">")[0])
                        sheet['A'+str(k)] = lastcamp
                        sheet['B'+str(k)] = found
                        try:
                            sheet['C'+str(k)] = lis[1].split(">")[0]
                        except:
                            print("..")
                        try:
                            l1 = found.split("<")[1].split(">")[0]
                            sheet['D'+str(k)] = l1
                        except:
                            print('00')
                        k = k+1
                        print("######################")
                        
                print("Over---------------------------------")
                    #kk = messages[i]
                
                    #print(colored(kk,'blue'))
                    
                    #print(colored(messages[i].body,'blue'))
        
        except:
            #print("www",messages[i].SenderEmailAddress)
            print(colored("Sender's Type Unknown: ",'red'),messages[i].Sender)
            #print('Class: ',messages[i].Sender)
work.save('C:/Users/Lumenci 3/Desktop/'+dt+key+add+'.xlsx')