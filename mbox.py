import re
import base64
import uuid
import time
import socket
from time import gmtime, strftime
from pymongo import MongoClient

Attachment_DIRECTORY = 'D:/Email'  # Attachment file store path

def Process_MBOX():
    nameoffile = input("Enter Path [Ex. D:/Email/mxbox] : ")
    host_name = socket.gethostname()
    host_ip = socket.gethostbyname(host_name)
    attachmenturl = ''
    file = open(nameoffile,encoding = "ISO-8859-1")
    f = file.read()
    file.close()
    emails = f.split('From ')
    print(len(emails))
    timestamp = int(time.time())

    for j in range(1,len(emails)):
        print(j)
        email_sender_id = ''
        email_recipient_id =''
        email_recipient_CC_id =''
        subject =''
        Date =''
        emailbody = ''
        email_recipient_CCO_ID =''

        email_sender_id1 = emails[j].split('\n', 1)[0]
        email_sender = str(email_sender_id1).split(' ')
        email_sender_id = email_sender[0]

        print(email_sender_id)
        #print(email_sender)

        if 'Delivered-To:' in emails[j]:
            deliverd = emails[j].split('Delivered-To:', 1)[1]
            deliverd = deliverd.split('\n', 1)[0]
            deliverd = re.sub("(\<.*?\>)", "", deliverd.strip())
            deliverd = re.sub('[ \t]+', ' ', deliverd.strip())
            deliverd = deliverd.replace('&nbsp;', '')
            deliverd = deliverd.strip()
            #print(deliverd)
        if 'To: ' in emails[j]:
            To = emails[j].split('To: ', 1)[1]
            To = To.split('\n', 1)[0]
            To = re.sub("(\<.*?\>)", "", To.strip())
            To = re.sub('[ \t]+', ' ', To.strip())
            To = To.replace('&nbsp;', '')
            email_recipient_id = To.strip()
            #print(To)
        if 'CC: ' in emails[j]:
            CC = emails[j].split('CC: ', 1)[1]
            CC = CC.split('\n', 1)[0]
            CC = re.sub("(\<.*?\>)", "", CC.strip())
            CC = re.sub('[ \t]+', ' ', CC.strip())
            CC = CC.replace('&nbsp;', '')
            email_recipient_CC_id = CC.strip()
            #print(CC)

        if 'Bcc: ' in emails[j]:
            Bcc = emails[j].split('Bcc: ', 1)[1]
            Bcc = Bcc.split('\n', 1)[0]
            Bcc = re.sub("(\<.*?\>)", "", Bcc.strip())
            Bcc = re.sub('[ \t]+', ' ', Bcc.strip())
            Bcc = Bcc.replace('&nbsp;', '')
            email_recipient_CCO_ID = Bcc.strip()
            #print(Bcc)

        if 'Subject: ' in emails[j]:
            Subject = emails[j].split('Subject: ', 1)[1]
            Subject = Subject.split('\n', 1)[0]
            Subject = re.sub("(\<.*?\>)", "", Subject.strip())
            Subject = re.sub('[ \t]+', ' ', Subject.strip())
            Subject = Subject.replace('&nbsp;', '')
            subject = Subject.strip()
            #print(Subject)

        if 'Date: ' in emails[j]:
            Date = emails[j].split('Date: ', 1)[1]
            Date = Date.split('\n', 1)[0]
            Date = re.sub("(\<.*?\>)", "", Date.strip())
            Date = re.sub('[ \t]+', ' ', Date.strip())
            Date = Date.replace('&nbsp;', '')
            Date = Date.strip()
            #print(Date)

        conn = ''
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

        if 'Content-Type: text/plain;' in emails[j]:
            text_plain = emails[j].split('Content-Type: text/plain;', 1)[1]
            text_plain = text_plain.split('Content-Type: ', 1)[0]
            text_plain = re.sub("(\<.*?\>)", "", text_plain.strip())
            text_plain = re.sub('[ \t]+', ' ', text_plain.strip())
            text_plain = text_plain.replace('&nbsp;', '')
            emailbody = text_plain.strip()
            # print(text_plain)

        if 'Content-Type: text/html;' in emails[j]:
            text_html = emails[j].split('Content-Type: text/html;', 1)[1]
            text_html = text_html.split('Content-Type: ', 1)[0]
            #print(text_html)
            html = text_html.split('\n\n', 1)[1]
            html = html.split('\n\n', 1)[0]
            name = str(uuid.uuid1()) + '.html'
            f = open('%s/%s' % (Attachment_DIRECTORY, name),'w')
            f.write(html)
            f.close()
            if attachmenturl == '':
                attachmenturl = Attachment_DIRECTORY + "/" + name
            else:
                attachmenturl = attachmenturl + "," + Attachment_DIRECTORY + "/" + name
            #text_html = text_html.strip()
            #print(text_html)

        if 'Content-Type: application/pdf;' in emails[j]:
            pqr = emails[j].split('Content-Type: application/pdf;')
            #print(len(pqr))
            for p in range(1, len(pqr)):
                name=pqr[p].split('name="',1)[1]
                name = name.split('"',1)[0]
                name = name.replace("..",".")
                name = str(uuid.uuid1()) + name
                pdf2 = pqr[p].split('\n\n', 1)[1]
                pdfbase = pdf2.split('\n\n', 1)[0]
                if 'Content-Type: text/html' in pdfbase:
                    pdfbase = pdfbase.split('--Apple-Mail',1)[0]
                #text_html = text_html.strip()
                with open('%s/%s' % (Attachment_DIRECTORY, name), 'wb') as theFile:
                    theFile.write(base64.b64decode(pdfbase))
                if attachmenturl == '':
                    attachmenturl = Attachment_DIRECTORY + "/" + name
                else:
                    attachmenturl = attachmenturl + "," + Attachment_DIRECTORY + "/" + name

        if 'Content-Type: image/jpeg' in emails[j]:
            name = ''
            abc = emails[j].split('Content-Type: image/jpeg')
            #print(len(abc))
            #print(abc[1])
            for c in range(1,len(abc)):
                if 'name="' in abc[c]:
                    name = abc[c].split('name="', 1)[1]
                    name = name.split('"', 1)[0]
                    name = name.replace("..", ".")
                    name = str(uuid.uuid1()) + name
                else:
                    name1 = abc[c].split('Content-ID:',1)[1]
                    name1 = name1.split('@',1)[0]
                    name = name1.split('.',1)[1]
                    name = str(name).replace('.','')
                    name = name+'.jpeg'
                    name = str(uuid.uuid1()) + name
                image = abc[c].split('\n\n', 1)[1]
                finalimage=image.split('\n\n',1)[0]
                #print(finalimage)
                with open('%s/%s' % (Attachment_DIRECTORY, name), 'wb') as theFile:
                    theFile.write(base64.b64decode(finalimage))

                if attachmenturl == '':
                    attachmenturl = Attachment_DIRECTORY + "/" + name
                else:
                    attachmenturl = attachmenturl + "," + Attachment_DIRECTORY + "/" + name

            emp_rec1 = {
                "tphashobject_metadata_tib": "8f7074d8-a520-4f7d-b2d3-09dc36acb5fd",
                "tphashobject_metadata_tib_name": "TPEMAIL",
                "tpemail_metadata_mail_box_name": "Frederico Mbox",
                "tpemail_metadata_id_mail_box": str(uuid.uuid1()),
                "tpemail_metadata_time": timestamp,
                "tpemail_metadata_time_zone": strftime("%z", gmtime()),
                "tpemail_metadata_email_subject": subject,
                "tpemail_metadata_email_sender": "",
                "tpemail_metadata_email_sender_id": email_sender_id,
                "tpemail_metadata_email_recipeint": "",
                "tpemail_metadata_email_recipient_id": email_recipient_id,
                "tpemail_metadata_email_timestamp": Date,
                "tpemail_metadata_email_header": "",
                "tpemail_metadata_email_body": emailbody,
                "tpemail_metadata_email_seq": "",
                "tpemail_metadata_email_text_content": "",
                "tpemail_metadata_email_html_content": "",
                "tpemail_metadata_email_eml_content": "",
                "tpemail_metadata_email_links": "",
                "tpemail_metadata_email_atach": attachmenturl,
                "tpemail_metadata_email_template_id": "",
                "tpemail_metadata_email_track_link": "",
                "tpemail_metadata_email_recipeint_cc": "",
                "tpemail_metadata_email_recipient_cc_id": email_recipient_CC_id,
                "tpemail_metadata_email_recipeint_cco": "",
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


Process_MBOX()
