from oauth2client.service_account import ServiceAccountCredentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError
from googleapiclient.discovery import build
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from typing import List
from constant import *
import mimetypes
import os.path
import smtplib
import base64
import time
import ssl
import gc
from dotenv import load_dotenv

load_dotenv()
app_pw = os.getenv("app_pw") # GCP app password


def goog_auth(SCOPES, SERVICE_ACCOUNT_KEY_FILE): # ,CLIENT_FILE):
    creds = None
    ''' For non-service accounts
    if os.path.exists(payroll_path + '\\Data\\token.json'):
        creds = Credentials.from_authorized_user_file(payroll_path + '\\Data\\token.json', SCOPES)
    '''
    # For service account
    creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_KEY_FILE, SCOPES)
    # If there are no (valid) credentials available, let the user log in.creds.refresh_tok
    if not creds:  # or not creds.valid:
        pass
        try:
            if creds and creds.expired and en:
                try:
                    creds.refresh(Request())
                except Exception as e:
                    print(e)
                    flow = InstalledAppFlow.from_client_secrets_file(CLIENT_FILE, SCOPES)
                    creds = flow.run_local_server(port=0)
            else:
                flow = InstalledAppFlow.from_client_secrets_file(CLIENT_FILE, SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open(payroll_path + '\\Data\\token.json', 'w') as token:
                token.write(creds.to_json())
        except Exception as e:
            err = str(e)
            if err == '(access_denied) ':
                raise Exception('Access has been denied by the user!')
    return creds


def create_file(service, file_name, file_type, parent_folder_id=None, upload_file=None):
    media = ''
    if file_type == 'spreadsheet':
        mimeType = 'application/vnd.google-apps.spreadsheet'
    elif file_type == 'folder':
        mimeType = 'application/vnd.google-apps.folder'
    elif file_type == 'text':
        file_metadata = {"name": file_name, 'parents': [parent_folder_id]}
        media = MediaFileUpload(upload_file, mimetype="text/plain")
    try:
        if parent_folder_id is None and file_type != 'text':
            file_metadata = {'name': file_name, 'mimeType': mimeType,}
        elif file_type != 'text':
            file_metadata = {'name': file_name, 'parents': [parent_folder_id], 'mimeType': mimeType,}

        # pylint: disable=maybe-no-member
        file = service.files().create(body=file_metadata, media_body=media, fields="id").execute()
        return file.get("id")

    except HttpError as error:
        print(f"An error occurred: {error}")
        return None


def get_file_id(service, file_name, file_type, parent_folder_id=None):
    if file_type == 'spreadsheet':
        mimeType = 'application/vnd.google-apps.spreadsheet'
    elif file_type == 'folder':
        mimeType = 'application/vnd.google-apps.folder'
    page_token = None
    if parent_folder_id is None:
        q = "name='{}', mimeType='{}'".format(file_name, mimeType)
    else:
        q = "name='{}' and parents='{}' and mimeType='{}'".format(file_name, parent_folder_id, mimeType)
    while True:
        # pylint: disable=maybe-no-member
        response = (
            service.files()
            .list(
                q=q,
                spaces="drive",
                fields="nextPageToken, files(id, name)",
                pageToken=page_token,
            )
            .execute()
        )
        for file in response.get("files", []):
            print(f'Found file: {file.get("name")}, {file.get("id")}')
            return file.get("id")
        page_token = response.get("nextPageToken", None)
        if page_token is None:
            break
    return None


def get_spreadsheet_data(service, spreadsheet_id):
    # Call the Sheets API
    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id, includeGridData=True).execute()
    val = spreadsheet['sheets']
    del spreadsheet
    gc.collect()
    return val


def get_sheet_values(service, spreadsheetId, range):
    return service.spreadsheets().values().get(spreadsheetId=spreadsheetId, range=range).execute()


def update_spreadsheet(service, val, spreadsheet_id, sheet_range):
    body = {"values": val}
    spreadsheet = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=sheet_range, body=body, valueInputOption="USER_ENTERED").execute()


class GmailException(Exception):
    """gmail base exception class"""


class NoEmailFound(GmailException):
    """no email found"""


def search_emails(service, query_string: str, label_ids: List = None):
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
        print(e)
        raise NoEmailFound('No emails returned')


def get_file_data(service, message_id, attachment_id):
    response = service.users().messages().attachments().get(
        userId='me',
        messageId=message_id,
        id=attachment_id
    ).execute()

    file_data = base64.urlsafe_b64decode(response.get('data').encode('UTF-8'))
    return file_data


def get_message_detail(service, message_id, msg_format='metadata', metadata_headers: List = None):
    message_detail = service.users().messages().get(
        userId='me',
        id=message_id,
        format=msg_format,
        metadataHeaders=metadata_headers
    ).execute()
    return message_detail


def send_email(sender_email, to, subject, body='', file_attachments=None, cc=None, bcc=None):
    mimeMessage = MIMEMultipart()
    mimeMessage['from'] = company_name
    mimeMessage['to'] = to
    mimeMessage['cc'] = cc
    mimeMessage['bcc'] = bcc
    mimeMessage['subject'] = subject
    mimeMessage.attach(MIMEText(body, 'plain'))

    # Attach files
    for attachment in file_attachments:
        content_type, encoding = mimetypes.guess_type(attachment)
        main_type, sub_type = content_type.split('/', 1)
        file_name = os.path.basename(attachment)

        f = open(attachment, 'rb')
        myFile = MIMEBase(main_type, sub_type)
        myFile.set_payload(f.read())
        myFile.add_header('Content-Disposition', 'attachment', filename=file_name)
        encoders.encode_base64(myFile)
        f.close()

        mimeMessage.attach(myFile)

    text = mimeMessage.as_string()
    context = ssl.create_default_context()

    # or port 465
    # await aiosmtplib.send(mimeMessage, hostname="smtp.gmail.com", sender=company_name, port=587, username=sender_email, password=app_pw, start_tls=True)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, app_pw)
        server.sendmail(sender_email, to, text)


def get_labels(service):
    return service.users().labels().list(
        userId='me').execute()


def update_email_label(service, message_id, remove_label, add_label):
    msg_labels = {'removeLabelIds': [remove_label], 'addLabelIds': [add_label]}
    response = service.users().messages().modify(
        userId='me',
        id=message_id,
        body=msg_labels).execute()
