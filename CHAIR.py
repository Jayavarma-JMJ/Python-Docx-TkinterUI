from docx import Document
from tkinter import *
import tkinter as tk
import smtplib
from email.message import EmailMessage
from pathlib import Path
import csv

document = Document('demo.docx') ## Assinging the template document
table = document.tables ##Getting the list of tables as an index. The desired tabkle is at [0]

Frommail =  ""
sslport=""
tlsport = ""
smtp_server = ""
password = ""
sender_email = ""
receiver_email = []
Tomail = []
Ccmail = []

with open('config.csv') as confgfile:
    reader = csv.reader(confgfile)
    next(reader)
    for types,value in reader:
        if types == "smtp":
            smtp_server = value
            print(smtp_server)
        if types == "sslport":
            sslport = value
        if types == "tlsport":
            tlsport = value
        if types == "sender":
            Frommail = value
            sender_email = Frommail
        if types == "password":
            password = value
            #print(type(password))
        if types == "to":
            Tomail.append(value)
        if types == "cc":
            Ccmail.append(value)
            receiver_email = Tomail + Ccmail
        
def Mailsender():
    msg = EmailMessage()
    msg['From'] = Frommail
    msg['To'] = ", ".join(Tomail)
    msg['Cc'] = ", ".join(Ccmail)
    msg['Subject'] = str(HospitalNameEntry.get()) + "- CHAIR escalation - Reg"
    msg.set_content("Escalation of " + str(HospitalNameEntry.get()) + " for "
                     + str(ReasonEntry.get("1.0", "end-1c")) + ". Kindly find the attachments below")
    attachment_path = str(HospitalNameEntry.get()) +  str(DtofRefEntry.get()) + "CHAIR.docx"

    with open(attachment_path, 'rb') as f:
        file_data = f.read()
        file_name = Path(attachment_path).name
        msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)
        
    ## Tkinter TypeError if as_string() is not used. Tk expects String or Bytes-like object.
    server.sendmail(sender_email, receiver_email, msg.as_string())
    Mail_Result.config(text="Mail Sent!")
    msg = EmailMessage()
    
def quitconnection():
    server.quit()
    Mail_Result.config(text="Disconnected. Please reconnect to send mail")

def inittlsconnection():
    global server
    server = smtplib.SMTP(smtp_server, tlsport,timeout=3)
    server.set_debuglevel(2)
    server.starttls()
    server.login(sender_email, password)
    Mail_Result.config(text="Connected!")

    mailquitbutton = tk.Button(root, text="Quit connection", command=quitconnection) 
    mailquitbutton.grid(row=16, column=1, padx=5, pady=(5,10))

def initsslconnection():
    global server
    server = smtplib.SMTP_SSL(smtp_server, sslport,timeout=3)
    server.set_debuglevel(2)
    server.login(sender_email, password)
    Mail_Result.config(text="Connected!")

    mailquitbutton = tk.Button(root, text="Quit connection", command=quitconnection) 
    mailquitbutton.grid(row=16, column=1, padx=5, pady=(5,10))
      
def create_wordfile():
    table[0].cell(0,7).text =  str(HospitalNameEntry.get())
    table[0].cell(1,1).text =  str(HospitalCdEntry.get())
    table[0].cell(1,7).text =  str(HospitalNetEntry.get())
    table[0].cell(2,1).text =  str(HospitalCityEntry.get())
    table[0].cell(2,7).text =  str(ZoneEntry.get())
    table[0].cell(3,7).text =  str(ReasonEntry.get("1.0", "end-1c"))
    table[0].cell(4,7).text =  str(DtofRefEntry.get())
    table[0].cell(5,2).text =  str(ClaimNoEntry.get("1.0", "end-1c"))
    table[0].cell(5,7).text =  str(ClaimCntEntry.get())
    table[0].cell(6,7).text =  str(ClaimNoEntry.get("1.0", "end-1c"))
    table[0].cell(7,7).text =  str(IssueDetailEntry.get("1.0", "end-1c"))
    table[0].cell(8,7).text =  str(RmdActEntry.get("1.0", "end-1c"))

##    Following cells have already been filled in the excel
##    table[0].cell(9,1).text =
##    table[0].cell(10,1).text =
##    table[0].cell(10,7).text =
##    table[0].cell(11,7).text =
    
    FileName = str(HospitalNameEntry.get()) +  str(DtofRefEntry.get()) + "CHAIR.docx"
    document.save(FileName)
    output_label.config(text="Word Document Generated")
    Docname.config(text=FileName)
    mail_button = tk.Button(root, text="Send Mail", command=Mailsender) 
    mail_button.grid(row=16, column=2, padx=5, pady=(5,10))

def reset_fields():
    for widget in root.winfo_children():
        if isinstance(widget, tk.Entry):  # Check if the widget is an Entry field
            widget.delete(0, tk.END)
        if isinstance(widget, tk.Text):  # Check if the widget is an Text field
            widget.delete('1.0', tk.END)

##    del msg['Subject']
    output_label.config(text="Text has been reset")
    Mail_Result.config(text="")
    Docname.config(text="")

#Setting Up GUI using TKinter    

root = tk.Tk()
root.title("CHAIR Escalation")

HospitalName = tk.Label(root, text="Enter Hospital Name:")
HospitalName.grid(row=0, column=0, padx=5, pady=0) 
HospitalNameEntry = tk.Entry(root, width=30)
HospitalNameEntry.grid(row=1, column=0, padx=5, pady=(0,5))  

HospitalCd = tk.Label(root, text="Enter Hospital Code:")
HospitalCd.grid(row=0, column=1, padx=5, pady=0) 
HospitalCdEntry = tk.Entry(root, width=20)
HospitalCdEntry.grid(row=1, column=1, padx=5, pady=(0,5)) 

HospitalNet = tk.Label(root, text=" NW or NNW ?")
HospitalNet.grid(row=0, column=2, padx=5, pady=0) 
HospitalNetEntry = tk.Entry(root, width=5)
HospitalNetEntry.grid(row=1, column=2, padx=5, pady=(0,5)) 

HospitalCity = tk.Label(root, text="Enter Hospital City:")
HospitalCity.grid(row=2, column=0, padx=5, pady=0) 
HospitalCityEntry = tk.Entry(root, width=20)
HospitalCityEntry.grid(row=3, column=0, padx=5, pady=(0,5)) 

Zone = tk.Label(root, text="Zone:")
Zone.grid(row=2, column=1, padx=5, pady=0) 
ZoneEntry = tk.Entry(root, width=20)
ZoneEntry.grid(row=3, column=1, padx=5, pady=(0,5))

DtofRef = tk.Label(root, text="Date of Referrence:")
DtofRef.grid(row=2, column=2, padx=5, pady=0) 
DtofRefEntry = tk.Entry(root, width=20)
DtofRefEntry.grid(row=3, column=2, padx=5, pady=(0,5)) 

Reason = tk.Label(root, text="Enter Brief Reason for Referral:")
Reason.grid(row=4, column=0, padx=5, pady=0, columnspan=2) 
ReasonEntry = tk.Text(root, width=50, height=3)
ReasonEntry.grid(row=5, column=0, columnspan=3, padx=5, pady=(0,5)) 

ClaimNo = tk.Label(root, text="Enter Claim numbers (seperated by spaces/enter):") 
ClaimNo.grid(row=8, column=0, padx=5, pady=0, columnspan=2)
ClaimNoEntry = tk.Text(root, width=25, height=2)
ClaimNoEntry.grid(row=9, column=0, padx=5, pady=(0,5), columnspan=2) 

ClaimCnt = tk.Label(root, text=" No. of claims:")
ClaimCnt.grid(row=8, column=2, padx=5, pady=0) 
ClaimCntEntry = tk.Entry(root, width=5)
ClaimCntEntry.grid(row=9, column=2, padx=5, pady=(0,5)) 

IssueDetail = tk.Label(root, text=" Detailed explanation:")
IssueDetail.grid(row=10, column=0, padx=5, pady=0) 
IssueDetailEntry = tk.Text(root, width=50, height=5)
IssueDetailEntry.grid(row=11, column=0, padx=5, pady=(0,5), columnspan=3) 

RmdAct = tk.Label(root, text="Recommended Action:")
RmdAct.grid(row=12, column=0, padx=5, pady=0) 
RmdActEntry = tk.Text(root, width=50, height=2)
RmdActEntry.grid(row=13, column=0, padx=5, pady=(0,5), columnspan=3)

## Document Name
Docfield = tk.Label(root, text= "File Name:")
Docfield.grid(row=14, column=0, padx=5, pady = (5,10))

Docname = tk.Label(root, text= "" )
Docname.grid(row=14, column=1, padx=5, pady=(5,10), columnspan=2)

## Submit Button
submit_button = tk.Button(root, text="Submit", command=create_wordfile) 
submit_button.grid(row=15, column=0, padx=5, pady=(5,10))

## Document status
output_label = tk.Label(root, text="")
output_label.grid(row=15, column=1, padx=5, pady=(5,10))

resetbutton = tk.Button(root, text="Reset Data", command=reset_fields)
resetbutton.grid(row=15, column=2, padx=5, pady=(5,10))

## Mail Button will be generated only after Word document is prepared.
## This is to prevent Emails without attachements.
## mail_button = tk.Button(root, text="Send Mail", command=Mailsender) 
## mail_button.grid(row=16, column=2, padx=5, pady=(5,10))

mailconnbutton = tk.Button(root, text="Connect(TLS)", command=inittlsconnection) 
mailconnbutton.grid(row=16, column=0, padx=5, pady=(5,10))

mailconnbutton = tk.Button(root, text="Connect(SSL)", command=initsslconnection) 
mailconnbutton.grid(row=17, column=0, padx=5, pady=(5,10))

##  Status updates related to mailing
Mail_Result = tk.Label(root, text="")
Mail_Result.grid(row=17, column=1, padx=5, pady=(5,10), columnspan=3)
 
root.mainloop()



          

