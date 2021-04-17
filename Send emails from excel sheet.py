# -*- coding: utf-8 -*-
"""
Created on Fri Apr 16 00:21:24 2021

@author: Prithwis
"""
print("FOR THIS APPLICATION TO RUN PROPERLY, YOU NEED TO GIVE ACCESS TO LESS SECURE APPS IN YOUR GMAIL ACCOUNT. VISIT: https://www.google.com/settings/security/lesssecureapps FOR ALLOWING.")
import xlrd
import pandas as pd
import smtplib
a = input("ENTER THE EMAIL ADDRESS FROM WHICH YOU WANT TO SEND THE EMAILS:")
sender_email = a
b = input("ENTER THE PASSWORD:")
password = b
import socket
socket.getaddrinfo('localhost', 8080)
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(sender_email, password)
excel_file = input("ENTER YOUR EXCEL SHEET FILE PATH:")
list_of_emails = pd.read_excel(excel_file)
email_column = input("ENTER THE COLUMN NAME OF THE COLUMN CONTAINING THE EMAIL ADDRESS:")
name_column = input("ENTER THE COLUMN NAME OF THE COLUMN CONTAINING THE NAMES OF THE RECIPIENTS:")
names_list = list_of_emails[name_column]
email_list = list_of_emails[email_column]
for i in range(len(email_list)):
    name = names_list[i]
    email = email_list[i]
    message = """ write message here"""
    server.sendmail(sender_email, [email], message)
server.close()    
print("MESSAGE SUCCESSFULLY SENT.")
