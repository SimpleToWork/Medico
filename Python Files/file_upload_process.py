import datetime
import os
import pandas as pd
from google_drive_class import GoogleDriveAPI
from google_sheets_api import GoogleSheetsAPI
from gmail_api import GoogleGmailAPI
from global_modules import print_color
import time


def get_form_data(x):
    GsheetAPI = GoogleSheetsAPI(credentials_file=x.gsheet_credentials_file, token_file=x.gsheet_token_file, scopes=x.gsheet_scopes,
                    sheet_id=x.google_sheet_form_responses)

    df = GsheetAPI.get_data_from_sheet(sheetname='Form responses 1', range_name='A:K')
    print_color(df, color='g')

    records_to_recruit = df[(df['processed'] != "TRUE")]
    if records_to_recruit.shape[0] >0:
        records_to_recruit['patient_dob'] = pd.to_datetime(records_to_recruit['patient_dob'], format="%d/%m/%Y")
        records_to_recruit['date_of_exam'] = pd.to_datetime(records_to_recruit['date_of_exam'] , format="%d/%m/%Y")
        records_to_recruit['timestamp'] = pd.to_datetime(records_to_recruit['timestamp'], format="%d/%m/%Y %H:%M:%S")
    print_color(records_to_recruit, color='y')

    return records_to_recruit


def process_records(x, records_to_recruit):
    GdriveAPI = GoogleDriveAPI(credentials_file=x.drive_credentials_file, token_file=x.drive_token_file,
                               scopes=x.drive_scopes)

    GsheetAPI = GoogleSheetsAPI(credentials_file=x.gsheet_credentials_file, token_file=x.gsheet_token_file,
                                  scopes=x.gsheet_scopes, sheet_id=x.google_sheet_response_detail)
    GsheetAPI_1 = GoogleSheetsAPI(credentials_file=x.gsheet_credentials_file, token_file=x.gsheet_token_file,
                                scopes=x.gsheet_scopes, sheet_id=x.google_sheet_form_responses)
    folder_name = "Uploads"

    folder_id = x.gmail_upload_folder_id
    print_color(folder_id, color='b')
    # folders = GdriveAPI.get_drive_folder(folder_name=folder_name)

    child_folders = GdriveAPI.get_child_folders(folder_id=folder_id)
    print_color(child_folders, color='b')

    response_folder_id = None


    for each_folder in child_folders:
        if each_folder.get("name") == 'RECORD-INPUT':
            response_folder_id = each_folder.get("id")
            break

    print_color(response_folder_id, color='y')
    sub_response_folders = GdriveAPI.get_child_folders(folder_id=response_folder_id)
    folder_dict = {x.get('name'): x.get("id") for x in sub_response_folders}
    folder_names = [x.get('name') for x in sub_response_folders]

    print_color(folder_dict, color='g')


    row_count = GsheetAPI.get_row_count(sheetname="Detailed Evaluation Data")

    for i in range(records_to_recruit.shape[0]):
        id = records_to_recruit['id'].iloc[i]
        files_to_upload = records_to_recruit['please_upload_up_to_10_files_here'].iloc[i]
        date_of_exam = records_to_recruit['date_of_exam'].iloc[i]
        print_color(date_of_exam, color='g')
        if str(date_of_exam) != 'NaT':
            file_name_date_of_exam = date_of_exam.strftime('%Y.%m.%d')
            date_of_exam = date_of_exam.strftime('%Y-%m-%d')
        else:
            file_name_date_of_exam = (datetime.datetime.now() + datetime.timedelta(days=365)).strftime('%Y.%m.%d')
            date_of_exam = (datetime.datetime.now() + datetime.timedelta(days=365)).strftime('%Y-%m-%d')
        patient_first_name = records_to_recruit['patient_first_name'].iloc[i]
        patient_last_name = records_to_recruit['patient_last_name'].iloc[i]
        timestamp = records_to_recruit['timestamp'].iloc[i].strftime('%Y-%m-%d %H:%M:%S')
        lawyer_name = records_to_recruit['lawyer_name'].iloc[i]
        patient_dob = records_to_recruit['patient_dob'].iloc[i].strftime('%Y-%m-%d')
        email_address = records_to_recruit['email_address'].iloc[i]

        name_of_new_folder = f'{file_name_date_of_exam} {patient_last_name}, {patient_first_name} '
        if name_of_new_folder not in folder_names:
            new_folder_id = GdriveAPI.create_folder(name_of_new_folder, response_folder_id)
        else:
            new_folder_id = folder_dict.get(name_of_new_folder)

        print_color(name_of_new_folder, new_folder_id, color='b')
        print_color(files_to_upload, color='y')

        unique_files = files_to_upload.split(",")
        unique_file_ids = [x.split("https://drive.google.com/open?id=")[-1] for x in unique_files]
        print_color(unique_file_ids, color='r')
        print_color(f'File to Move: {len(unique_file_ids)}', color='b')

        row_count += 1
        counter = -1
        data_to_upload = []
        for each_id in unique_file_ids:
            file_name = f"https://drive.google.com/open?id={each_id}"
            print_color(file_name, color='y')
            folder_name = f"https://drive.google.com/drive/u/0/folders/{new_folder_id}"
            data_to_upload.append([timestamp, id, each_id, file_name, lawyer_name, email_address, patient_first_name,
                                   patient_last_name, patient_dob, date_of_exam, folder_name])
            counter +=1
            GdriveAPI.move_file(file_id=each_id,new_folder_id=new_folder_id)

            print_color(data_to_upload, color='b')



        print_color(f'File count {len(unique_file_ids)}  Range {row_count} - {row_count+counter}', color='y')
        GsheetAPI.insert_row_to_sheet(sheetname="Detailed Evaluation Data", gid=0,
                            insert_range=['B', row_count, 'J', row_count+counter],
                            data=data_to_upload
                            )
            # break
        row_count += counter

        processed_responses = [[id, 'TRUE']]
        GsheetAPI_1.insert_row_to_sheet(sheetname="Processed Responses", gid=865982653,
                                      insert_range=['A', 1, 'B', 1],
                                      data=processed_responses
                                      )

        time.sleep(5)

        # break
    #


def run_file_upload_process(x, environment):
    records_to_recruit = get_form_data(x)
    process_records(x, records_to_recruit)