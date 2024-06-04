import datetime
import pandas as pd
from google_drive_class import GoogleDriveAPI
from google_sheets_api import GoogleSheetsAPI
from gmail_api import GoogleGmailAPI
from global_modules import print_color, error_handler, Get_SQL_Types, Change_Sql_Column_Types, engine_setup
from email_process import is_date
import time
from dateutil.parser import parse

def get_form_data(x):
    GsheetAPI = GoogleSheetsAPI(credentials_file=x.gsheet_credentials_file, token_file=x.gsheet_token_file, scopes=x.gsheet_scopes,
                    sheet_id=x.google_sheet_form_responses)

    df = GsheetAPI.get_data_from_sheet(sheetname='Form responses 1', range_name='A:K')
    print_color(df, color='g')

    records_to_recruit = df[(df['processed'] != "TRUE")]
    print_color(records_to_recruit, color='y')
    if records_to_recruit.shape[0] >0:

        records_to_recruit['patient_dob'] = pd.to_datetime(records_to_recruit['patient_dob'], format="%d/%m/%Y")
        # records_to_recruit['date_of_exam'] = pd.to_datetime(records_to_recruit['date_of_exam'] , format="%d/%m/%Y", errors='coerce')
        records_to_recruit['timestamp'] = pd.to_datetime(records_to_recruit['timestamp'], format="%d/%m/%Y %H:%M:%S")
    print_color(records_to_recruit, color='y')

    return records_to_recruit


def process_individual_record(records_to_recruit, i, folder_names, GdriveAPI, folder_dict, date, response_folder_id, row_count, GsheetAPI, GsheetAPI_1):
    id = records_to_recruit['id'].iloc[i]
    files_to_upload = records_to_recruit['please_upload_up_to_10_files_here'].iloc[i]
    patient_first_name = records_to_recruit['patient_first_name'].iloc[i]
    patient_last_name = records_to_recruit['patient_last_name'].iloc[i]
    timestamp = records_to_recruit['timestamp'].iloc[i].strftime('%Y-%m-%d %H:%M:%S')
    lawyer_name = records_to_recruit['lawyer_name'].iloc[i]
    patient_dob = records_to_recruit['patient_dob'].iloc[i].strftime('%Y-%m-%d')
    email_address = records_to_recruit['email_address'].iloc[i]
    date_of_exam = records_to_recruit['date_of_exam'].iloc[i]
    date_of_exam = pd.to_datetime(date_of_exam, errors='ignore', format="%d/%m/%Y")

    print_color(f'patient: {patient_first_name} {patient_last_name}, date_of_exam: {date_of_exam} ', color='r')
    date_processed = False
    if str(date_of_exam) != 'NaT':
        # print_color(is_date(date_of_exam, fuzzy=False))
        if isinstance(date_of_exam, pd.Timestamp):
            date_details = parse(str(date_of_exam), fuzzy=False)

            print_color(date_details, color='g')
            print_color(date_details.year, color='y')
            if date_details.year < 2000:
                print_color(f'Year is formatted incorrectly', color='r')
                name_of_new_folder = f'{date_of_exam}, {patient_last_name}, {patient_first_name} '
                return row_count, name_of_new_folder, date_processed
            else:
                if is_date(str(date_of_exam), fuzzy=False):
                    file_name_date_of_exam = date_of_exam.strftime('%Y.%m.%d')
                    date_of_exam = date_of_exam.strftime('%Y-%m-%d')
                    date_processed = True
                else:
                    name_of_new_folder = f'{date_of_exam}, {patient_last_name}, {patient_first_name} '
                    return row_count, name_of_new_folder, date_processed
        else:
            name_of_new_folder = f'{date_of_exam}, {patient_last_name}, {patient_first_name} '
            return row_count, name_of_new_folder, date_processed

    else:
        file_name_date_of_exam = (datetime.datetime.now() + datetime.timedelta(days=365)).strftime('%Y.%m.%d')
        date_of_exam = (datetime.datetime.now() + datetime.timedelta(days=365)).strftime('%Y-%m-%d')
        date_processed = True




    print_color(f'date_of_exam: {date_of_exam}', date_processed, color='g')


    if datetime.datetime.strptime(date_of_exam, "%Y-%m-%d") < datetime.datetime.strptime(date, "%Y-%m-%d"):
        print_color(f'Date is earlier than today', color='y')

        name_of_new_folder = f'{file_name_date_of_exam}, {patient_last_name}, {patient_first_name} MAIL'
    else:
        name_of_new_folder = f'{file_name_date_of_exam}, {patient_last_name}, {patient_first_name} '

    if date_processed is True:
        if name_of_new_folder not in folder_names:
            new_folder_id = GdriveAPI.create_folder(name_of_new_folder, response_folder_id)
        else:
            new_folder_id = folder_dict.get(name_of_new_folder)

        print_color(f'name_of_new_folder: {name_of_new_folder}, {new_folder_id}', color='b')
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
            counter += 1
            GdriveAPI.move_file(file_id=each_id, new_folder_id=new_folder_id)

            print_color(data_to_upload, color='b')
        file_move_status = []
        for each_id in unique_file_ids:
            data = GdriveAPI.get_file_data(file_id=each_id)
            parent_folder_id = data.get("parents")[0]
            if new_folder_id == parent_folder_id:
                file_move_status.append(True)
            else:
                file_move_status.append(False)

        all_processed = False
        if False in file_move_status:
            print_color(f'File Has Not Moved from Inputs', color='r')
            all_processed = False
        else:
            all_processed = True

        # print_color(data.get("owners")[0].get("emailAddress"), data.get("parents")[0], color='r',
        #             output_file=main_log_file)

        if all_processed is True:
            print_color(f'File count {len(unique_file_ids)}  Range {row_count} - {row_count + counter}', color='y')

            GsheetAPI.insert_row_to_sheet(sheetname="Detailed Evaluation Data", gid=0,
                                          insert_range=['B', row_count, 'J', row_count + counter],
                                          data=data_to_upload
                                          )

            row_count += counter

            processed_responses = [[id, 'TRUE']]
            GsheetAPI_1.insert_row_to_sheet(sheetname="Processed Responses", gid=865982653,
                                            insert_range=['A', 1, 'B', 1],
                                            data=processed_responses
                                            )

    return row_count, name_of_new_folder, date_processed

def process_records(x, records_to_recruit):

    engine= engine_setup(project_name=x.project_name, hostname=x.hostname, username=x.username, password=x.password,
                            port=x.port)

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

    print_color(f'response_folder_id: {response_folder_id}', color='y')
    sub_response_folders = GdriveAPI.get_child_folders(folder_id=response_folder_id)
    folder_dict = {x.get('name'): x.get("id") for x in sub_response_folders}
    folder_names = [x.get('name') for x in sub_response_folders]

    print_color(folder_dict, color='g')


    row_count = GsheetAPI.get_row_count(sheetname="Detailed Evaluation Data")
    date = datetime.datetime.now().strftime('%Y-%m-%d')
    date_now = datetime.datetime.now().strftime("%Y-%m-%d")
    table_name = 'program_performance'
    for i in range(records_to_recruit.shape[0]):
        now = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")

        try:
            row_count, name_of_new_folder, date_processed = process_individual_record(records_to_recruit, i, folder_names, GdriveAPI, folder_dict, date,
                                  response_folder_id, row_count, GsheetAPI, GsheetAPI_1)

            if date_processed is True:
                executed = True
            else:
                executed = False
        except Exception as e:
            print_color(f'An Error Occurred', color='r')
            print_color(e, color='r')
            executed = False

        if executed is True:
            print_color(name_of_new_folder, date_processed, color='b')
            performance_list = [None, "Upload Process", date_now, now, name_of_new_folder, executed]
            performance_df = pd.DataFrame([performance_list])
            performance_df.columns = ['id', 'module_name', 'date', 'datetime', 'Patient_Folder__Name', 'module_complete']
            print_color(performance_df, color='g')
            sql_types = Get_SQL_Types(performance_df).data_types
            Change_Sql_Column_Types(engine=engine, Project_name=x.project_name, Table_Name=table_name, DataTypes=sql_types,
                                    DataFrame=performance_df)
            performance_df.to_sql(name=table_name, con=engine, if_exists='append', index=False, schema=x.project_name,
                                  chunksize=1000, dtype=sql_types)

            print_color(f'Data imported to {table_name}', color='g')

        # break



def run_file_upload_process(x, environment):
    GmailAPi = GoogleGmailAPI(credentials_file=x.gmail_credentials_file, token_file=x.gmail_token_file, scopes=x.gmail_scopes)
    # try:
    records_to_recruit = get_form_data(x)
    process_records(x, records_to_recruit)

    # except Exception as error:
    #     now = datetime.datetime.now().strftime("%Y-%m-%d")
    #     error_message = f"{type(error).__name__}, {error}"
    #
    #     email_body = \
    #     f'''Hello,
    #     <br><br>An Error Occurred on The Upload Process.
    #     <br>See Error Below
    #     <br><span style="color:Red;font-weight:Bold; ">{str(error_message)}</span>
    #
    #     <br><br>Thank you,
    #     <br><br>This is an automatically generated email.
    #     '''
    #     print_color(error_message, color='y')  # An error occurred: NameError
    #     GmailAPi.send_email(email_to=", ".join(x.notification_email), email_sender=x.email_sender,
    #                         email_subject=f'Upload Process Error {now}', email_cc=None, email_bcc=None,
    #                         email_body=email_body)