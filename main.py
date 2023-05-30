from keep_alive import keep_alive

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
from PyPDF2 import PdfReader
import base64
from typing import List
import time
from google_apis import create_service
import re

Sheet = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

bula = ServiceAccountCredentials.from_json_keyfile_name('credentials.json')

file= gspread.authorize(bula)
workbook = file.open("logistics")
sheet = workbook.sheet1


class GmailException(Exception):
    """gmail base exception class"""


class NoEmailFound(GmailException):
    """no email found"""

def search_emails(query_stirng: str, label_ids: List = None):
    try:
        message_list_response = service.users().messages().list(
            userId='me',
            labelIds=label_ids,
            q=query_string
        ).execute()

        message_items = message_list_response.get('messages')
        next_page_token = message_list_response.get('nextPageToken')

        while next_page_token:
            message_list_response = service.users().messages().list(
                userId='me',
                labelIds=label_ids,
                q=query_string,
                pageToken=next_page_token
            ).execute()

            message_items.extend(message_list_response.get('messages'))
            next_page_token = message_list_response.get('nextPageToken')
        return message_items
    except Exception as e:
        raise NoEmailFound('No emails returned')


def get_file_data(message_id, attachment_id, file_name, save_location):
    response = service.users().messages().attachments().get(
        userId='me',
        messageId=message_id,
        id=attachment_id
    ).execute()

    file_data = base64.urlsafe_b64decode(response.get('data').encode('UTF-8'))
    return file_data


def get_message_detail(message_id, msg_format='metadata', metadata_headers: List = None):
    message_detail = service.users().messages().get(
        userId='me',
        id=message_id,
        format=msg_format,
        metadataHeaders=metadata_headers
    ).execute()
    return message_detail

CLIENT_FILE = 'client-secret.json'
API_NAME = 'gmail'
API_VERSION = 'v1'
SCOPES = ['https://mail.google.com/']
service = create_service(CLIENT_FILE, API_NAME, API_VERSION, SCOPES)

query_string = 'logistics@nayaraenergy.com is:unread'
save_location = os.getcwd()

keep_alive()
while True:
    email_messages = search_emails(query_string)
    test_list = []

    if email_messages == None:
        continue

    elif email_messages != None:
        for email_message in email_messages:
            messageDetail = get_message_detail(email_message['id'], msg_format='full', metadata_headers=['parts'])
            messageDetailPayload = messageDetail.get('payload')

            body = messageDetailPayload.get('body')
            if not body:
                parts = messageDetailPayload.get('parts')
                if parts:
                    data = parts[0]['body']
                    body = data.get('data')
            else:
                body = body.get('data')

            if body:
                text = base64.urlsafe_b64decode(body).decode('utf-8')
                text = re.sub(r'<.*?>', '', text)  # Remove HTML tags
                text = re.sub(" ", '', text)  # Remove HTML tags
                gmailtxt = str(text)

                # For Date
                srch4 = re.search('10%EthanolBlendedMotorSpirit', gmailtxt)
                srch5 = re.search('AutomotiveDieselFuelBSVI', gmailtxt)
                if (srch4!=None):
                    index = srch4.end()
                    tarik = gmailtxt[int(index) + 15:int(index) + 25]
                if (srch5!=None):
                    index = srch5.end()
                    tarik = gmailtxt[int(index) + 8:int(index) + 18]
                test_list.append(tarik)

                # For Gadi
                srch3 = re.search('MP13H[0-9][0-9][0-9][0-9]', gmailtxt)
                gadi = srch3.group()
                test_list.append(gadi)

                # For Petrol
                srch1 = re.search('10%EthanolBlendedMotorSpirit', gmailtxt)
                p1 = ""
                if (srch1!=None):
                    index = srch1.end()
                    check = int(gmailtxt[int(index)+35:int(index)+37])
                    print(check)
                    if (check==50 or check==40 or check==80 or check==90):
                        p1 = gmailtxt[int(index)+43:int(index)+46]
                        p1 = re.sub("[a-zA-Z]", '', p1)
                        print(p1)
                    elif (check==12 or check==13 or check==14 or check==15 or check==17 or check==18 or check==20 or check==22):
                        p1 = gmailtxt[int(index) + 45:int(index) + 48]
                        p1 = re.sub("[a-zA-Z]", '', p1)
                        print(p1)
                    else:
                        p1 = gmailtxt[int(index) + 44:int(index) + 47]
                        p1 = re.sub("[a-zA-Z]", '', p1)
                        print(p1)
                test_list.append(p1)

                # For Diesel
                srch2 = re.search('AutomotiveDieselFuelBSVI', gmailtxt)
                d1 = ""
                if (srch2!=None):
                    index1 = srch2.end()
                    check1 = int(gmailtxt[int(index1) + 28:int(index1) + 30])
                    print(check1)
                    if (check1 == 50 or check1 == 40 or check1 == 80 or check1 == 90):
                        d1 = gmailtxt[int(index1) + 36:int(index1) + 39]
                        d1 = re.sub("[a-zA-Z]", '', d1)
                        print(d1)
                    elif (check1 == 12 or check1 == 13 or check1 == 14 or check1 == 15 or check1 == 17 or check1 == 18 or check1 == 20 or check1 == 22):
                        d1 = gmailtxt[int(index1) + 38:int(index1) + 41]
                        d1 = re.sub("[a-zA-Z]", '', d1)
                        print(d1)
                    else:
                        d1 = gmailtxt[int(index1) + 37:int(index1) + 40]
                        d1 = re.sub("[a-zA-Z]", '', d1)
                        print(d1)
                test_list.append(d1)

                sheet.append_row(test_list, table_range="A1:D1")
                test_list.clear()
                service.users().messages().modify(userId='me', id=email_message['id'], body={
                    'removeLabelIds': ['UNREAD']
                }).execute()
            else:
                print(f'No text found for email')
