from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload
import io

import os
from global_modules import print_color


class GoogleDriveAPI():
    def __init__(self, credentials_file=None, token_file=None, scopes=None):
        self.scopes =scopes
        self.credentials_file = credentials_file
        self.token_file = token_file
        self.inJsonFile = ''
        self.outFile = ''
        self.authenticate()

    def authenticate(self):
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists( self.token_file):
            creds = Credentials.from_authorized_user_file( self.token_file,  self.scopes )
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.credentials_file,  self.scopes )
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open( self.token_file, 'w') as token:
                token.write(creds.to_json())

        self.service = build('drive', 'v3', credentials=creds)
        # print_color( self.service)

    def get_drive_folder(self, folder_name):
        folders = []
        page_token = None
        while True:
            response = self.service.files().list(q=f"mimeType='application/vnd.google-apps.folder' and name = '{folder_name}'",
                                            spaces='drive',
                                            fields='nextPageToken, '
                                                   'files(id, name)',
                                            pageToken=page_token).execute()
            for folder in response.get('files', []):
                # Process change
                print(F'Found file: {folder.get("name")}, {folder.get("id")}')
            folders.extend(response.get('files', []))
            page_token = response.get('nextPageToken', None)
            if page_token is None:
                break

        print_color(folders, color='g')
        return folders

    def get_child_folders(self, folder_id=None):
        results = []
        page_token = None

        while True:
            response = self.service.files().list(
                q=f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.folder'",
                fields='nextPageToken, files(id, name)',
                pageToken=page_token
            ).execute()

            for file in response.get('files', []):
                results.append({'id': file['id'], 'name': file['name']})

            page_token = response.get('nextPageToken')
            if not page_token:
                break

        return results

    def get_files(self, folder_id):
        results = []
        page_token = None

        while True:
            response = self.service.files().list(
                q=f"'{folder_id}' in parents and mimeType!='application/vnd.google-apps.folder'",
                fields='nextPageToken, files(id, name)',
                pageToken=page_token
            ).execute()

            for file in response.get('files', []):
                results.append({'id': file['id'], 'name': file['name']})

            page_token = response.get('nextPageToken')
            if not page_token:
                break

        return results

    def download_file(self, file_id, file_name):
        try:
            request = self.service.files().get_media(fileId=file_id)
            file = open(file_name, 'wb')
            downloader = MediaIoBaseDownload(file, request)
            done = False
            while done is False:
                status, done = downloader.next_chunk()
                print(F'Download {int(status.progress() * 100)}.')

        except HttpError as error:
            print(F'An error occurred: {error}')
            file = None

        print('File downloaded to:', file_name)
        # return file.getvalue()

        # print_color(type(file.getvalue()), color='r')
        # return file

    def move_file(self, file_id, new_folder_id):
        print_color(file_id, color='y')
        file = self.service.files().get(fileId=file_id, fields='parents').execute()
        previous_parents = ",".join(file.get('parents'))

        if new_folder_id != previous_parents:

            # Move the file to the new folder
            file = self.service.files().update(
                fileId=file_id,
                addParents=new_folder_id,
                removeParents=previous_parents,
                fields='id, parents'
            ).execute()

            print_color(f'File Moved from {previous_parents} To {new_folder_id}')
        else:
            print_color(f'New Folder and Parent Folder are the same', color='r')

    def create_folder(self, folder_name, parent_folder):
        """ Create a folder and prints the folder ID
        Returns : Folder Id

        Load pre-authorized user credentials from the environment.
        TODO(developer) - See https://developers.google.com/identity
        for guides on implementing OAuth2 for the application.
        """


        try:
            # create drive api client
            file_metadata = {
                'name': folder_name,
                'mimeType': 'application/vnd.google-apps.folder',
                'parents': [parent_folder]
            }

            # pylint: disable=maybe-no-member
            file = self.service.files().create(body=file_metadata, fields='id').execute()
            print(F'Folder ID: "{file.get("id")}".')
            return file.get('id')

        except HttpError as error:
            print(F'An error occurred: {error}')
            return None
