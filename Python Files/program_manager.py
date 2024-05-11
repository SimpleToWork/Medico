from merge_files import merge_files_to_pdf, ocr_conversion
from global_modules import ProgramCredentials, print_color, record_program_performance, record_program_details
from email_process import run_email_process
from file_upload_process import run_file_upload_process
from email_diagnostic import run_email_diagnostic
from google_sheets_api import google_sheet_update
from google_drive_class import GoogleDriveAPI
import getpass
from tqdm import tqdm
import random
import time

import sys



def run_program(environment, function_to_run):
    x = ProgramCredentials(environment)

    if function_to_run == 'Email Process':
        run_email_process(x, environment)
    elif function_to_run == 'Upload Process':
        run_file_upload_process(x, environment)
    elif function_to_run == 'Merge Files':
        merge_files_to_pdf(x, environment)
    elif function_to_run == 'Email Diagnostic':
        run_email_diagnostic(x, environment)

    record_program_details(x, program_name="Medico", method=function_to_run)
    record_program_performance(x, program_name="Medico", method=function_to_run)



if __name__ == '__main__':
    if len(sys.argv) == 1:
        environment = 'production'
        function_to_run = 'Upload Process'
    else:
        environment = sys.argv[1]
        function_to_run = sys.argv[2]

    wait_time = random.randint(3, 35)
    computer = getpass.getuser()

    # if computer != "Ricky":
    #     print_color(f'Waiting {wait_time} seconds to start', color='y')
    #     for i in tqdm(range(wait_time)):
    #         time.sleep(1)

    run_program(environment, function_to_run)
