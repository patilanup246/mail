import datetime
import email
import email.header
import imaplib
import mimetypes
import socket
import time
import uuid
from time import gmtime, strftime
import re
import base64


import re
#from pymongo import MongoClient
from PST import main
# from test import Process_MBOX

Attachment_DIRECTORY = 'D:/Email'  # Attachment file store path
EMAIL_FOLDER = "INBOX"  # Gmail scrap folder name
host_name = socket.gethostname()
host_ip = socket.gethostbyname(host_name)



def List_Service(cond):
    print("List of services :")
    print("1. Emails Service")
    print("2. MBOX service (mac)")
    print("3. PST Service (Outlook)")

    SERVICENUMBER = int(input("Enter service number [Ex. 1] : "))

    if SERVICENUMBER == 1:
        print("Selected Service is Emails")
        from gmail import process_mailbox
        return False
    elif SERVICENUMBER == 2:
        print("Selected Service is MBOX")
        from mbox import Process_MBOX
        #Process_MBOX()
        return False
    elif SERVICENUMBER == 3:
        print("Selected Service is PST")
        nameoffile = input("Enter Path of PST File [Ex. D:/Email/backup_diego.pst] : ")
        main(nameoffile, 'PST Report')
        return False
    else:
        print("Please select correct service !!")
        return True


cond =True;
while cond:
    cond= List_Service(cond)

