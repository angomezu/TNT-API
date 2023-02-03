import json
import base64
import os
import pickle
import gspread
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials

# Load client secrets
CLIENT_SECRET = json.loads('{"installed":{"client_id":"963632109139-07maeo5l6tko3lhomdk7of7uvkoqgl67.apps.googleusercontent.com","project_id":"tnt-api-376720","auth_uri":"https://accounts.google.com/o/oauth2/auth","token_uri":"https://oauth2.googleapis.com/token","auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs","client_secret":"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","redirect_uris":["http://localhost"]}}')

# Authenticate to Gmail
creds = None
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)
if not creds or not creds.valid:
    creds = Credentials.from_authorized_user_info(info=CLIENT_SECRET)
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

# Connect to Gmail API
service = build('gmail', 'v1', credentials=creds)

# Find the email with the desired subject
results = service.users().messages().list(userId='me', q='subject:Registration Form').execute()

# Get the contents of the email
if results:
    msg = service.users().messages().get(userId='me', id=results['messages'][0]['id']).execute()
    payload = msg['payload']
    headers = payload['headers']
    body = payload['body']
    data = body['data']
    data = data.replace("-","+").replace("_","/")
    data = base64.b64decode(data)

# Connect to Google Sheets
gc = gspread.authorize(creds)

# Open the desired sheet
sh = gc.open_by_url("https://docs.google.com/spreadsheets/d/141eYzW24gX6gOnvPLBbWJzAJzKVek6AEzWJZ2_9Jxmc/edit#gid=522629157")
worksheet = sh.worksheet("TEST")

# Write the contents of the email to the sheet
worksheet.append_row([data])