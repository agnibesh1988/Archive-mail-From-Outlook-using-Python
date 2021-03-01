# -*- coding: utf-8 -*-
"""
Created on Wed Jul  8 15:06:45 2020

@author: Agnibesh.Samanta
"""



import win32com.client
import re
import os




               
if __name__ == '__main__':
    # loaction of the directory where emails need to save
    path = os.path.join("C:\\","mail_backup_box")


    if not os.path.exists(path):
        os.mkdir(path)
        print("Directory " , path ,  " Created ")
    else:    
        print("Directory " , path ,  " already exists")
        
        
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    accounts= win32com.client.Dispatch("Outlook.Application").Session.Accounts
    for account in accounts:
        inbox = outlook.Folders(account.DeliveryStore.DisplayName)
        folders = inbox.Folders
        for folder in folders:
            if str(folder) == "Inbox":
                messages = folder.Items
                for message in messages:
                    sender_name = message.SenderEmailAddress.split('-')[-1]
                    name = str(message.subject)
                    name = re.sub('[^A-Za-z0-9]+', ' ', name)+'.msg'
                    print(name)
                
                    #save email
                    message.SaveAs(path+'//'+'{}_{}_{}'.format(message.senton.date(),sender_name,name))
