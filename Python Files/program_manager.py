from merge_files import merge_files_to_pdf
from google_drive_class import GoogleDriveAPI
from global_modules import ProgramCredentials, print_color
from email_process import run_email_process
from file_upload_process import run_file_upload_process
from google_sheets_api import google_sheet_update
import getpass
import platform
import datetime
import pandas as pd





def run_program(environment, function_to_run):
    x = ProgramCredentials(environment)
    # merge_files_to_pdf()
    if function_to_run == 'Email Process':
        run_email_process(x, environment)
    elif function_to_run == 'Upload Process':
        run_file_upload_process(x, environment)

    google_sheet_update(x, program_name="Medico", method=function_to_run)


if __name__ == '__main__':

    environment = 'development'
    function_to_run = 'Upload Process'

    run_program(environment, function_to_run)
