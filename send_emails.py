import os.path
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io
import pandas as pd
from googleapiclient.errors import HttpError
import time

# Define SCOPES for the APIs
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/drive.readonly'
]

def authenticate():
    """Authenticate and create service objects for Sheets, Gmail, and Drive APIs."""
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('sheets', 'v4', credentials=creds), build('gmail', 'v1', credentials=creds), build('drive', 'v3', credentials=creds)

def download_file(service, file_id):
    """Download file from Google Drive."""
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    fh.seek(0)
    return fh

def send_emails(sheet_service, gmail_service, drive_service, spreadsheet_id, file_id):
    """Send emails with the attachment."""
    # Read spreadsheet data
    sheet = sheet_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range='Sheet1').execute()
    values = sheet.get('values', [])
    
    # Iterate over each email entry in the spreadsheet
    for row in values[1:]:  # Skip header row
        company_name = row[0]  # Name of the company
        email_address = row[1]  # Recipient email address
        
        # Define the subject and body of the email
        subject = ''
        body = ''
        
        # Download the file anew for each email
        file_content = download_file(drive_service, file_id)
        
        # Create the email message
        message = MIMEMultipart()
        message['To'] = email_address
        message['From'] = 'me'  # The authenticated sender's email
        message['Subject'] = subject
        
        # Attach the email body
        message.attach(MIMEText(body, 'plain'))
        
        # Create the attachment
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(file_content.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="resume.pdf"')
        
        # Attach the file to the message
        message.attach(part)
        
        # Encode the message and send it
        raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
        
        try:
            gmail_service.users().messages().send(userId='me', body={'raw': raw_message}).execute()
            print(f'Sent email to {email_address}')
        except HttpError as error:
            print(f'An error occurred: {error}')
        
        # Delay to avoid hitting rate limits
        time.sleep(5)  # Adjust the sleep time as needed

if __name__ == '__main__':
    # Replace with your actual spreadsheet ID and file ID
    SPREADSHEET_ID = 'your_spreadsheet_id_here'
    FILE_ID = 'your_file_id_here'
    
    # Authenticate and get the API services
    sheets_service, gmail_service, drive_service = authenticate()
    
    # Send emails using the data from the spreadsheet and the specified file attachment
    send_emails(sheets_service, gmail_service, drive_service, SPREADSHEET_ID, FILE_ID)
