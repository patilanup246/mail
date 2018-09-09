import os
import sys
import argparse
import logging
import re
import pypff
import socket
import time
import uuid
from time import gmtime, strftime
from pymongo import MongoClient


output_directory = ""
date_list = ''
# date_dict = {x:0 for x in xrange(1, 25)}
# date_list = [date_dict.copy() for x in xrange(7)]


def main(pst_file, report_name):
    """
    The main function opens a PST and calls functions to parse and report data from the PST
    :param pst_file: A string representing the path to the PST file to analyze
    :param report_name: Name of the report title (if supplied by the user)
    :return: None
    """
    logging.debug("Opening PST for processing...")
    pst_name = os.path.split(pst_file)[1]
    opst = pypff.open(pst_file)
    root = opst.get_root_folder()

    logging.debug("Starting traverse of PST structure...")
    folderTraverse(root)

    logging.debug("Generating Reports...")



def folderTraverse(base):
    """
    The folderTraverse function walks through the base of the folder and scans for sub-folders and messages
    :param base: Base folder to scan for new items within the folder.
    :return: None
    """
    for folder in base.sub_folders:
        #print(folder.name)
        if folder.number_of_sub_folders:
            folderTraverse(folder) # Call new folder to traverse:
        checkForMessages(folder)


def checkForMessages(folder):
    """
    The checkForMessages function reads folder messages if present and passes them to the report function
    :param folder: pypff.Folder object
    :return: None
    """
    logging.debug("Processing Folder: " + folder.name)
    message_list = []
    #print(folder.sub_messages)
    for message in folder.sub_messages:
        #print(message.subject)
        message_dict = processMessage(message)
        message_list.append(message_dict)
    folderReport(message_list, folder.name)


def processMessage(message):
    """
    The processMessage function processes multi-field messages to simplify collection of information
    :param message: pypff.Message object
    :return: A dictionary with message fields (values) and their data (keys)
    """
    #print(message)
    abc = message.transport_headers
    #print(str(abc))
    if 'From: ' in str(abc):
        deliverd = str(abc).split('From: ', 1)[1]
        deliverd = deliverd.split('\n', 1)[0]

    email_sender = ''
    email_sender_id = ''
    try:
        if '<' in str(deliverd):
            email_sender = str(deliverd).split('<')[0]
            email_sender_id = str(deliverd).split('<')[1].replace(">", " ")
        else:
            email_sender_id = deliverd
    except:
        print("")
    print(email_sender)
    print(message.subject)

    if 'To: ' in str(abc):
        To = str(abc).split('To: ', 1)[1]
        To = To.split('\n', 1)[0]

    email_recipeint = ''
    email_recipient_id = ''
    try:
        if '<' in str(To):
            email_recipeint = str(To).split('<')[0]
            email_recipient_id = str(To).split('<')[1].replace(">", " ")
        else:
            email_recipient_id = str(To)
    except:
        print("")

    if 'Cc: ' in str(abc):
        CC = str(abc).split('Cc: ', 1)[1]
        CC = CC.split('\n', 1)[0]
        # CC if not found
    email_recipeint_CC = ''
    email_recipient_CC_id = ''
    try:
        if '<' in str(CC):
            email_recipeint_CC = str(CC).split('<')[0]
            email_recipient_CC_id = str(CC).split('<')[1].replace(">", " ")
        else:
            email_recipient_CC_id = str(CC)
    except:
        print("")

    if 'Bcc: ' in str(abc):
        Bcc = str(abc).split('Bcc: ', 1)[1]
        Bcc = Bcc.split('\n', 1)[0]
    email_recipeint_CCO = ''
    email_recipient_CCO_ID = ''
    try:
        if '<' in str(Bcc):
            email_recipeint_CCO = str(Bcc).split('<')[0]
            email_recipient_CCO_ID = str(Bcc).split('<')[1].replace(">", " ")
        else:
            email_recipient_CCO_ID = str(Bcc)
    except:
        print("")
    Date = ''
    if 'Date: ' in str(abc):
        Date = str(abc).split('Date: ', 1)[1]
        Date = Date.split('\n', 1)[0]
        Date = re.sub("(\<.*?\>)", "", Date.strip())
        Date = re.sub('[ \t]+', ' ', Date.strip())
        Date = Date.replace('&nbsp;', '')
        Date = Date.strip()

    return {
        "subject": message.subject,
        "sender": message.sender_name,
        "header": message.transport_headers,
        "body": message.plain_text_body,
        "email_sender":email_sender,
        "email_sender_id":email_sender_id,
        "email_recipeint":email_recipeint,
        "email_recipient_id":email_recipient_id,
        "email_recipeint_CC":email_recipeint_CC,
        "email_recipient_CC_id":email_recipient_CC_id,
        "email_recipeint_CCO":email_recipeint_CCO,
        "email_recipient_CCO_ID":email_recipient_CCO_ID,
        "Date":Date,
        "attachment_count": message.number_of_attachments,
    }


def folderReport(message_list, folder_name):
    """
    The folderReport function generates a report per PST folder
    :param message_list: A list of messages discovered during scans
    :folder_name: The name of an Outlook folder within a PST
    :return: None
    """
    if not len(message_list):
        logging.warning("Empty message not processed")
        return
    host_name = socket.gethostname()
    host_ip = socket.gethostbyname(host_name)
    timestamp = int(time.time())
    for m in message_list:
        if m['body']:
            emailbody = m['body']
            sender = m['sender']
            header = m['header']
            subject = m['subject']
            attachment = m['attachment_count']
            email_sender = m['email_sender']
            email_sender_id = m['email_sender_id']
            email_recipeint = m['email_recipeint']
            email_recipient_id = m['email_recipient_id']
            email_recipeint_CC = m['email_recipeint_CC']
            email_recipient_CC_id = m['email_recipient_CC_id']
            email_recipeint_CCO = m['email_recipeint_CCO']
            email_recipient_CCO_ID = m['email_recipient_CCO_ID']
            Date = m['Date']

            try:
                conn = MongoClient('localhost', 27017)
                print("Connected successfully!!!")
            except:
                print("Could not connect to MongoDB")

            # database
            db = conn.hash
            # Created or Switched to collection names: testGmailAnup
            collection = db.ib

            if db.ib.find({'email_timestamp': str(Date)}).count() > 0:
                continue

            emp_rec1 = {
                "tphashobject_metadata_tib": "8f7074d8-a520-4f7d-b2d3-09dc36acb5fd",
                "tphashobject_metadata_tib_name": "TPEMAIL",
                "tpemail_metadata_mail_box_name": "Frederico PST",
                "tpemail_metadata_id_mail_box": str(uuid.uuid1()),
                "tpemail_metadata_time": timestamp,
                "tpemail_metadata_time_zone": strftime("%z", gmtime()),
                "tpemail_metadata_email_subject": subject,
                "tpemail_metadata_email_sender": email_sender,
                "tpemail_metadata_email_sender_id": email_sender_id,
                "tpemail_metadata_email_recipeint": email_recipeint,
                "tpemail_metadata_email_recipient_id": email_recipient_id,
                "tpemail_metadata_email_timestamp": Date,
                "tpemail_metadata_email_header": "",
                "tpemail_metadata_email_body": emailbody,
                "tpemail_metadata_email_seq": "",
                "tpemail_metadata_email_text_content": "",
                "tpemail_metadata_email_html_content": "",
                "tpemail_metadata_email_eml_content": "",
                "tpemail_metadata_email_links": "",
                "tpemail_metadata_email_atach": attachment,
                "tpemail_metadata_email_template_id": "",
                "tpemail_metadata_email_track_link": "",
                "tpemail_metadata_email_recipeint_cc": email_recipeint_CC,
                "tpemail_metadata_email_recipient_cc_id": email_recipient_CC_id,
                "tpemail_metadata_email_recipeint_cco": email_recipeint_CCO,
                "tpemail_metadata_email_recipient_cco_id": email_recipient_CCO_ID,
                "tphashobject_metadata_hash_owner_id": "g00zNU6n7WfhUI1u4A5ebxSN0732",
                "tpemail_metadata_hash_sender_id": "",
                "tpemail_metadata_hash_recipt_id": "",
                "tpemail_metadata_hash_sender_name": "",
                "tpemail_metadata_hash_receipt_name": "",
                "tpemail_metadata_hash_recipt_cc_id": "",
                "tpemail_metadata_hash_recipt_cco_id": "",

                "tphashobject_metadata_hub_group_id": "da0a7b22-fb15-46e0-9f5a-019263d79e36",
                "tphashobject_metadata_data_sinc_mongodb": "",
                "tphashobject_metadata_action": "",
                "tphashobject_metadata_role": "",
                "tphashobject_metadata_layout_role": "",
                "tphashobject_metadata_group_id": "9529b03b-38e8-4bdc-aa62-8055a4c36a55",
                "tpemailbox_metadata_IP_machine": host_ip

            }

            # Insert Data
            rec_id1 = collection.insert_one(emp_rec1)
            print("Data inserted with record ids", rec_id1)