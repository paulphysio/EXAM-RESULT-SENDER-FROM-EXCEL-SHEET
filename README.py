# EXAM-RESULT-SENDER-FROM-EXCEL-SHEET
# This is a program that helps a University course adviser to be able to send the exam results of a particular students from a list of many students from an excel sheet
import smtplib # importing library
from email.message import EmailMessage
#from email.mime.text import MIMEText

port = 465 #server port

import xlrd as xl
import openpyxl as pl

file = xl.open_workbook(r"C:\Users\hp\Documents\TESTER.xlsx")#Opening of excel sheet
wb = pl.load_workbook(r"C:\Users\hp\Documents\TESTER.xlsx")#Loading of excel sheet

sheet_ad = file.sheet_by_index(0)#Accessing First sheet

sheet = wb["Sheet1"]#Accessing First sheet

for i in range(sheet_ad.nrows + 1):
    if i > 1:
        #message to be sent
        msg1 = sheet["A1"].value +"  |  " + str(sheet["A" + str(i)].value) + "\n" +sheet["B1"].value +"  |  " + str(sheet["B" + str(i)].value) + "\n" +sheet["C1"].value +"  |  " + str(sheet["C" + str(i)].value) + "\n" +sheet["D1"].value +"  |  " + str(sheet["D" + str(i)].value) + "\n" +sheet["E1"].value +"  |  " + str(sheet["E" + str(i)].value) + "\n" +sheet["F1"].value +"  |  " + str(sheet["F" + str(i)].value) + "\n" +sheet["G1"].value +"  |  " + str(sheet["G" + str(i)].value) + "\n" 
        #msg2 = str(sheet["A" + str(i)].value) +" | "+ str(sheet["B" + str(i)].value) + "  |  " + str(sheet["C" + str(i)].value) + "  |  " + str(sheet["D" + str(i)].value) + "  |  " + str(sheet["E" + str(i)].value) + "  |  " + str(sheet["F" + str(i)].value) + "  |  " + str(sheet["G" + str(i)].value) + "  |  " 
        #msg3 = "Sorry this is Mbappe testing my code"

        sender_email = "Emilyjohnson25099@gmail.com" #senders email
        password = "Physio350" #senders email 
        
        receiver_mail = sheet["G" + str(i)].value #receivers email
        
        
        
        
        with smtplib.SMTP_SSL("smtp.gmail.com", port) as smtp:
        
        
        
            
            smtp.login(sender_email, password) #login in
            email_message = EmailMessage()
            email_message["From"] = "\"The President of SPE\" <Emilyjohnson25099.com>";
            email_message["To"] = receiver_mail
            email_message["Subject"] = "SPE email tester"
            
            email_message.set_content(msg1 + "\n" + msg3)
            
            print("Login Succesful")
            smtp.send_message(email_message ) #sending message
            print("Message has been sent to " + receiver_mail)
