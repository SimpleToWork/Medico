import os.path
import sys
import os
currentdir = os.path.dirname(os.path.realpath(__file__))
parentdir = os.path.dirname(currentdir)
sys.path.append(parentdir)
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient import discovery, http
from global_modules import print_color
from email import utils, encoders
from email.mime.text import MIMEText
from email.message import EmailMessage
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime import application, multipart, text, base, image, audio
import mimetypes
import base64

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']


class GoogleGmailAPI():
    def __init__(self,  credentials_file=None, token_file=None, scopes=None):
        self.scopes = scopes
        self.credentials_file = credentials_file
        self.token_file = token_file
        self.authenticate()

    def authenticate(self):
        creds = None

        if os.path.exists(self.token_file):
            creds = Credentials.from_authorized_user_file(self.token_file, self.scopes)
            # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())

            else:
                flow = InstalledAppFlow.from_client_secrets_file(self.credentials_file, self.scopes)
                creds = flow.run_local_server(port=0)

            # Save the credentials for the next run
            with open(self.token_file, 'w') as token:
                token.write(creds.to_json())
        #     return creds
        # else:
        #     return creds

        self.service = build('gmail', 'v1', credentials=creds)

    def get_emails(self, query):
        """Shows basic usage of the Gmail API.
        Lists the user's Gmail labels.
        """

        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.

        try:
            # Call the Gmail API
            results = self.service.users().messages().list(userId='me',q=query).execute()
            print_color(results, color='r')
            # results = service.users().labels().list(userId='me').execute()
            messages = results.get('messages', [])

            if not messages:
                print('No messages found.')
                return
            print('messages:')
            message_list = []
            for message in messages:
                print(message['id'])
                message_result = self.service.users().messages().get(userId='me', id=message['id']).execute()
                message_list.append(message_result)
                # print_color(message_result, color='y')

            return message_list


        except HttpError as error:
            # TODO(developer) - Handle errors from gmail API.
            print(f'An error occurred: {error}')

    def get_file_attachments(self, message_id, attachment_id):
        attachment_results = self.service.users().messages().attachments().get(userId='me',messageId=message_id, id=attachment_id).execute()
        print_color(attachment_results, color='y')

        return attachment_results

    def attach_file(self, output_folder, file_name, message):
        filename = output_folder + "\\" + file_name
        # In same directory as script
        # TEST_NAME = "TEST NAME"
        # Open PDF file in binary mode
        with open(filename, "rb") as attachment:
            # Add file as application/octet-stream
            # Email client can usually download this automatically as attachment
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())

        # Encode file in ASCII characters to send by email
        encoders.encode_base64(part)

        # Add header as key/value pair to attachment part
        part.add_header('Content-Disposition', 'attachment', filename=file_name)
        # Add attachment to message and convert message to string
        message.attach(part)

        return message

    def send_email(self, email_to, email_sender, email_subject, email_cc, email_bcc, email_body, files=[]):

        message = MIMEMultipart()
        message["From"] = email_sender
        message["To"] = email_to
        message["Subject"] = email_subject
        message['date'] = utils.formatdate(localtime=True)
        message["Bcc"] = email_bcc
        message["Cc"] = email_cc
        # message["HTMLBody"] = body

        # Add body to email
        # message.attach(MIMEText(body, "plain"))

        message.attach(MIMEText(email_body, "html"))

        for each_file in files:
            export_folder = "\\".join(each_file.split("\\")[0:-1])
            file_name = each_file.split("\\")[-1]

            message = self.attach_file(export_folder, file_name, message)



        encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

        body = {
            'message': {
                'raw': encoded_message
            }
        }
        # message = (self.service.users().messages().send(userId='me', body=body).execute())
        try:
            message = (self.service.users().drafts().create(userId='me', body=body).execute())
            print_color(f'Message for {email_subject} Has Been Sent', color='g')
            return True
        except Exception as e:
            print_color(e, color='r')
            print_color(f'Email Could not be sent', color='r')
            return False