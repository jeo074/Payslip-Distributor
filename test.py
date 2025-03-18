from test_google import goog_auth, get_spreadsheet_data, get_sheet_values
from test_google import get_file_id, create_file, send_email, update_spreadsheet
from googleapiclient.discovery import build

CLIENT_FILE = r'C:\Users\jjdino\Python - Copy\Ops\Tool Resources\client_secret.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.file']
creds = goog_auth(SCOPES, CLIENT_FILE)

sheets_service = build("sheets", "v4", credentials=creds)
drive_service = build("drive", "v3", credentials=creds)

text_file = r'C:\Users\jjdino\Python - Copy\Ops\API test.txt'

id = create_file(drive_service, 'test api', 'text', '1Wrs8NgALitiuAdtwtDwDnDELzNYdbLcs', upload_file=text_file)
pass