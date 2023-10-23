from merge_files import merge_files_to_pdf
from google_drive_class import GoogleDriveAPI
from global_modules import ProgramCredentials, print_color
from email_process import run_email_process
from file_upload_process import run_file_upload_process
from google_sheets_api import GoogleSheetsAPI
import getpass
import platform
import datetime
import pandas as pd



def google_sheet_update(project_folder, program_name, method):
    client_secret_file = f'{project_folder}\\Text Files\\client_secret.json'
    token_file = f'{project_folder}\\Text Files\\token.json'
    sheet_id = '19FUWyywrtS4JTbOHW_GqDSEl0orqu99XCJJFa4upVlw'
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    df = GoogleSheetsAPI(client_secret_file,token_file, sheet_id, SCOPES).get_data_from_sheet(sheetname='KYF ACHWORKS', range_name='A:E')
    print_color(df, color='r')

    row_number = df.shape[0] + 2

    computer_name = platform.node()
    user = getpass.getuser()
    time_now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


    data_list = [time_now, computer_name, user, program_name, method, True]
    df = pd.DataFrame([data_list])
    print_color(df)
    # GoogleSheetsAPI(client_secret_file,token_file, sheet_id, SCOPES).write_data_to_sheet(data =df ,sheetname='KYF ACHWORKS',
    #   row_number=row_number,include_headers=False,  clear_data=False)


def run_program(environment, function_to_run):
    x = ProgramCredentials(environment)
    # merge_files_to_pdf()
    if function_to_run == 'email_process':
        run_email_process(x, environment)
    elif function_to_run == 'upload_process':
        run_file_upload_process(x, environment)



if __name__ == '__main__':
    environment = 'development'
    function_to_run = 'email_process'

    run_program(environment, function_to_run)
