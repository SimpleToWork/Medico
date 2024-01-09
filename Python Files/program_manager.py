from merge_files import merge_files_to_pdf
from google_drive_class import GoogleDriveAPI
from global_modules import ProgramCredentials, print_color, engine_setup
from email_process import run_email_process
from file_upload_process import run_file_upload_process
from google_sheets_api import google_sheet_update
import getpass
import platform
from tqdm import tqdm
import random
import time
import datetime
import pandas as pd

import sys



def run_program(environment, function_to_run):
    x = ProgramCredentials(environment)

    if function_to_run == 'Email Process':
        run_email_process(x, environment)
    elif function_to_run == 'Upload Process':
        run_file_upload_process(x, environment)
    elif function_to_run == 'Merge Files':
        merge_files_to_pdf(x, environment)

    computer = getpass.getuser()
    if computer != "Ricky":
        google_sheet_update(x, program_name="Medico", method=function_to_run)


if __name__ == '__main__':
    if len(sys.argv) == 1:
        environment = 'production'
        function_to_run = 'Upload Process'
    else:
        environment = sys.argv[1]
        function_to_run = sys.argv[2]

    wait_time = random.randint(3, 35)
    computer = getpass.getuser()

    if computer != "Ricky":
        print_color(f'Waiting {wait_time} seconds to start', color='y')
        for i in tqdm(range(wait_time)):
            time.sleep(1)

    run_program(environment, function_to_run)
