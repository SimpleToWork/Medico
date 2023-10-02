from __future__ import print_function
import os
import sys
currentdir = os.path.dirname(os.path.realpath(__file__))
parentdir = os.path.dirname(currentdir)
sys.path.append(parentdir)
import os.path
from sys import argv
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import pandas as pd
import numpy as np
# from sqlalchemy import inspect
# import global_modules
# from global_modules import engine_setup, print_color, run_sql_scripts


class GoogleSheetsAPI():
    def __init__(self, credentials_file=None, token_file=None, scopes=None, sheet_id=None):
        self.credentials_file = credentials_file
        self.token_file = token_file
        self.sheet_id = sheet_id
        self.scopes = scopes

        self.service = self.service_setup()


    def service_setup(self):
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists( self.token_file ):
            creds = Credentials.from_authorized_user_file( self.token_file , self.scopes)
            # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(self.credentials_file, self.scopes)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open( self.token_file , 'w') as token:
                token.write(creds.to_json())

        service = build('sheets', 'v4', credentials=creds)

        return service

    def create_new_spreadsheet(self, title):
        spreadsheet = {
            'properties': {
                'title': title
            }
        }
        spreadsheet = self.service.spreadsheets().create(body=spreadsheet,fields='spreadsheetId').execute()
        spreadsheet_id = spreadsheet.get('spreadsheetId')

        print('Spreadsheet ID: {0}'.format(spreadsheet.get('spreadsheetId')))


    def get_data_from_sheet(self, sheetname, range_name):
        """Shows basic usage of the Sheets API.
        Prints values from a sample spreadsheet.
        """


        # Call the Sheets API
        sheet = self.service.spreadsheets()
        result = sheet.values().get(spreadsheetId=self.sheet_id,
                                    range=f'{sheetname}!{range_name}').execute()
        values = result.get('values', [])
        df = pd.DataFrame(values)
        if df.shape[0] > 0:
            new_header = df.iloc[0]
            df.columns = [x.lower().replace(" ", "_") for x in new_header]
            df = df[1:]

        return df

    def write_data_to_sheet(self, data,sheetname, row_number, include_headers=True, clear_data=False):

        for col in data.columns:
            if data[col].dtype == float:
                pass
                # print(data[col].dtype)
            else:
                data[col] = data[col].astype(str)
        data = data.replace(np.nan, "")
        data = data.replace('None', "")

        if include_headers is True:
            column_values = [x.replace("_", " ").title() for x in data.columns]
        values = data.values.tolist()
        if include_headers is True:
            values.insert(0, column_values)
        # values = [['TEST', 'TEST 2',
        #            'Abc','Def']]
        # for i in range(0, len(values), 5000):
        #     upated_values = values[i:i+5000]
        body = {
            'values': values
        }
        #
        #     print(sheet_id)
        if clear_data is True:
            self.service.spreadsheets().values().clear(spreadsheetId=self.sheet_id, range="A1:AC1000000")

        result = self.service.spreadsheets().values().update(
            spreadsheetId=self.sheet_id, range=f'{sheetname}!A{row_number}',
            valueInputOption='USER_ENTERED', body=body).execute()

        print(f'Data Updated')


