import os

from constant import *
import json
from dotenv import load_dotenv

load_dotenv()
messenger_page_id = os.getenv("messenger_page_id")
page_access_token = os.getenv("page_access_token")


def get_psids(msgr, fb_accounts):
    total_accounts = len(fb_accounts)
    page_count = 1
    psids = {}
    url = 'https://graph.facebook.com/v{}/{}/conversations?fields=participants&access_token={}'.format(fb_api_version, messenger_page_id, page_access_token)
    while total_accounts != len(psids):
        response = msgr.get(url)
        response_json = response.json()
        response_text = str(response_json).lower()
        done = []
        for employee_name, fb_account in fb_accounts.items():
            name_str = "'{}'".format(fb_account).lower()
            if name_str in response_text:
                id_str = response_text.split(name_str, 1)[1].split("}", 1)[0]
                id_str = id_str.split("'id': '")[1].split("'")[0]
                psids[employee_name] = id_str
                done.append(employee_name)
        for accounts in done:
            del fb_accounts[accounts]
        if 'next' in response_json['paging']:
            url = response_json['paging']['next']
        if page_count == 20:
            break
        page_count += 1
    return psids


def send_message(msgr, psid, message):
    headers = {'content-type': 'application/json'}
    request_data = {
          'recipient': {'id': psid},
          'messaging_type': "MESSAGE_TAG",
          'tag': 'CONFIRMED_EVENT_UPDATE',
          'message': {'text': message},
        }
    url = 'https://graph.facebook.com/v{}/{}/messages?access_token={}'.format(fb_api_version, messenger_page_id, page_access_token)
    return msgr.post(url, data=json.dumps(request_data), headers=headers)


def send_attachment(msgr, psid, file_path):
    file_name = file_path.split('\\')[-1]
    file = {'filedata': (file_name, open(file_path, 'rb'), 'file/pdf')}
    request_data = {'recipient': (None, {'id': psid}), 'messaging_type': (None, 'MESSAGE_TAG'), 'tag':
                    (None, 'CONFIRMED_EVENT_UPDATE'), 'message': (None, {'attachment': {'type': 'file', 'payload': {}}})}
    url = 'https://graph.facebook.com/v{}/{}/messages?access_token={}'.format(fb_api_version, messenger_page_id, page_access_token)
    return msgr.post(url, data=request_data, files=file)
