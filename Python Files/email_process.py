import datetime

import pandas as pd
from dateutil.parser import parse
from google_drive_class import GoogleDriveAPI
from google_sheets_api import GoogleSheetsAPI
from gmail_api import GoogleGmailAPI
from global_modules import print_color, create_folder


from docx import Document
import docx2txt
import docx2pdf
from openpyxl.utils import get_column_letter
from dateutil.parser import parse
# import textract
from win32com import client as wc
import time
import io
import os
import getpass
import re

w = wc.Dispatch('Word.Application')

def is_date(string, fuzzy=False):
    """
    Return whether the string can be interpreted as a date.

    :param fuzzy: bool, ignore unknown tokens in string if True
    """
    try:
        parse(string, fuzzy=fuzzy)
        return True

    except ValueError:
        return False


def upload_data_to_gsheet(x, GsheetAPI, df):
    sheet_name = x.auto_publish_sheet_name

    # print_color(id_column, color='b')

    row_count = GsheetAPI.get_row_count(sheetname=sheet_name)
    print_color(row_count, color='g')

    data = df.values.tolist()

    print_color(data, color='b')
    end_column =get_column_letter(len(df.columns))
    print_color(end_column, color='r')

    dropdown_items = [
        {"userEnteredValue": 'TRUE',},
        {"userEnteredValue": 'FALSE',},
    ]

    GsheetAPI.insert_row_to_sheet(sheetname=sheet_name, gid=0,
                                  insert_range = ["A", 1, end_column, 1],
                                  data=data,
                                  # insert_dropdown=True, dropdown_values=dropdown_items,
                                  # dropdown_range=[10, 1, 11, 2]
                                  copy_area=True,
                                  copy_source_range=[0, 10, 2, 12, 3],
                                  copy_destinations_range=[0, 10, 1, 12, 2],
                                  copy_pasteType='PASTE_DATA_VALIDATION'
                                  )


def process_doc_file(x, GdriveAPI, GsheetAPI, file_id, file_name, id_number):
    create_folder(f'{x.project_folder}\\Extracts')
    file_export = f'{x.project_folder}\\Extracts\\{file_name}'

    file_type = file_name.split(".")[-1]


    GdriveAPI.download_file(file_id=file_id, file_name=file_export)

    data_dict = {
        "file_name": file_name,
        "date_of_report": None,
        "attorney_name": None,
        "attorney_email": None,
        "patient_name": None,
        "date_of_birth": None,
        "date_of_arrival": None,
        "all_fields_assigned": False,
        "approved_to_send_out": None,
    }

    first_line_text = 0
    date_of_report = None
    attorney_name = None


    # doc = Document(file_export)

    if file_type == 'doc':
        doc = w.Documents.Open(os.path.abspath(file_export))
        doc.SaveAs(file_export, 16)
        print_color(f'{file_export} Saved as Docx', color='r')

        # file_export = file_export.replace(".doc", ".docx")
        # file_type = file_export.split(".")[-1]
        text = docx2txt.process(file_export)

    elif file_type == 'docx':
        text = docx2txt.process(file_export)

        # text = textract.process(file_export)
    text_list = text.split("\n")
    # print_color(text_list, color='y')


    for i, row_text in enumerate(text_list):
        if i <= 30:
            text = row_text
            print_color(text, color='y')


            if first_line_text ==0:
                if text.strip() != "":
                    if is_date(text.strip()):
                        first_line_text += 1


            # lines.append(paragraph.text)
            # print(paragraph.text)
            if first_line_text != 0 and date_of_report is None:
                date_of_report = text

                data_dict["date_of_report"] = [date_of_report]
            if "RE:" in text.upper():
                data_dict["patient_name"] = [text.upper().replace("RE:","").replace(",","").replace("\t","").strip()]
            elif "NAME: " in text.upper():
                data_dict["patient_name"] = [text.upper().replace("NAME:", "").replace(",", "").replace("\t", "").strip()]
            if "DOB" in text.upper():
                data_dict["date_of_birth"] =[text.upper().replace("DOB","").replace(",","").replace("\t","").replace(":","").strip()]
            if "D.O.B" in text.upper():
                data_dict["date_of_birth"] = [text.upper().replace("D.O.B", "").replace(",", "").replace("\t", "").replace(":", "").strip()]
            if "DOA" in text.upper():
                data_dict["date_of_arrival"] = [text.upper().replace("DOA","").replace("\t","").replace(":","").replace(",","\n").strip()]
            if "DOI" in text.upper():
                data_dict["date_of_arrival"] = [text.upper().replace("DOA", "").replace("\t", "").replace(":", "").replace(",", "\n").strip()]
            if "ESQ" in text.upper():
                attorney_name = text.upper()
                data_dict["attorney_name"] = [attorney_name.replace("ESQ","").replace(".","").replace(",","").strip()]
            if attorney_name is None:

                if first_line_text <= i <= 5:
                    print_color(f'attorney_name Here {i}')
                    print_color(text, color='r')
                    if text.upper() != "":

                        attorney_name = text.upper()
                        data_dict["attorney_name"] = [ attorney_name.replace("ESQ", "").replace(".", "").replace(",", "").strip()]


            if "@" and ".com" in text:
                data_dict["attorney_email"] = [text]

    if data_dict.get('date_of_report') is not None and \
        data_dict.get('attorney_name') is not None and \
        data_dict.get('attorney_email') is not None and \
        data_dict.get('patient_name') is not None and \
        data_dict.get('date_of_birth') is not None and \
        data_dict.get('date_of_arrival') is not None:
        data_dict["all_fields_assigned"] = True
        data_dict["approved_to_send_out"] = True




    print_color(data_dict, color='r')

    df = pd.DataFrame.from_dict(data_dict ,orient ='columns')

    print_color(df, color='y')

    df.insert(0, "ID", id_number)
    df.insert(1, "Import Date", datetime.datetime.now().strftime('%Y-%m-%d'))
    df.insert(2, "File ID", file_id)

    upload_data_to_gsheet(x, GsheetAPI, df)


def convert_doc_to_pdf(x, file_name):
    file_export = f'{x.project_folder}\\Extracts\\{file_name}'
    file_type = file_name.split(".")[-1]

    pdf_export = f'{x.project_folder}\\Extracts\\{file_name.replace(file_type, "pdf")}'
    print_color(file_export, color='p')
    print_color(pdf_export, color='p')


    print_color( os.path.exists(file_export), color='g')
    if file_type == 'doc':
        doc = w.Documents.Open(os.path.abspath(file_export))
        doc.SaveAs(file_export, 16)
        print_color(f'{file_export} Saved as Docx', color='r')

    converted_to_pdf = False
    try:
        docx2pdf.convert(file_export, pdf_export)
        print_color(f'File Converted to PDF', color='g')
        converted_to_pdf = True
    except Exception as e:
        print_color(e, color='r')
        converted_to_pdf = False

    return converted_to_pdf, pdf_export, file_type


def is_date(string, fuzzy=False):
    """
    Return whether the string can be interpreted as a date.

    :param string: str, string to check for date
    :param fuzzy: bool, ignore unknown tokens in string if True
    """
    try:
        if string is not None:
            parse(string, fuzzy=fuzzy)

        return True

    except ValueError:
        return False


def email_doc_out(x, GmailAPI, file_name, date_of_report, attorney_name, attorney_email, patient_name, dob, doa, pdf_export):
    if  is_date(dob, fuzzy=False):
        dob = datetime.datetime.strptime(dob,"%Y-%m-%d").strftime('%m/%d/%Y')
    if is_date(doa, fuzzy=False):
        doa = datetime.datetime.strptime(doa,"%Y-%m-%d").strftime('%m/%d/%Y')
    if is_date(date_of_report, fuzzy=False):
        date_of_report = datetime.datetime.strptime(date_of_report, "%Y-%m-%d").strftime('%m/%d/%Y')


    email_body = \
    f'''The report on your client has been completed and is attaced.
        <br><span style="color:Black;font-weight:Bold; ">DOB:</span> {dob}
        <br><span style="color:Black;font-weight:Bold; ">DOA:</span> {doa}
        <br><span style="color:Black;font-weight:Bold; ">Date of Report: </span> {date_of_report}
    <br><br>Please contact our office if needed.

    <br><br>Medico-Legal Evaluations
    <br>732-972-4471
    '''

    print_color(pdf_export, color='g')
    subject = f"Report for: {patient_name} {date_of_report}"
    email_sent = GmailAPI.send_email(email_to=attorney_email,
                        email_sender=x.email_sender,
                        email_subject = subject,
                        email_cc = "",
                        email_bcc = "",
                        email_body=email_body,
                        files=[pdf_export])

    return email_sent


def update_google_sheet_record(GsheetAPI, id, file_id, file_converted, email_sent, file_moved, new_file_folder):
    sheetname = 'Converted Files'
    # row_number = GsheetAPI.get_row_count(sheetname) +1
    id_file_id = f'{id}{file_id}'
    data = [[id, file_id,id_file_id, file_converted, email_sent, file_moved, new_file_folder]]
    df = pd.DataFrame(data)

    GsheetAPI.insert_row_to_sheet(sheetname=sheetname, gid=1915603262,
                                  insert_range=['A', 1, "D", 1],
                                  data=None)

    row_number = 2
    GsheetAPI.write_data_to_sheet(df, sheetname, row_number, include_headers=False, clear_data=False)


def move_drive_file(GdriveAPI, file_name, file_id, child_folders):
    file_folder_dict = {
        'A': 'A-C',
        'B': 'A-C',
        'C': 'A-C',
        'D': 'D-F',
        'E': 'D-F',
        'F': 'D-F',
        'G': 'G-I',
        'H': 'G-I',
        'I': 'G-I',
        'J': 'J-L',
        'K': 'J-L',
        'L': 'J-L',
        'M': 'M-O',
        'N': 'M-O',
        'O': 'M-O',
        'P': 'P-R',
        'Q': 'P-R',
        'R': 'P-R',
        'S': 'S-U',
        'T': 'S-U',
        'U': 'S-U',
        'V': 'V-Z',
        'W': 'V-Z',
        'X': 'V-Z',
        'Y': 'V-Z',
        'Z': 'Z - Test Cases'

    }
    file_start_letter = file_name[0]
    move_file_folder = file_folder_dict.get(file_start_letter)
    move_folder_id = None

    for each_item in child_folders:

        if each_item['name'] == move_file_folder:
            move_folder_id = each_item['id']
            break

    print_color(file_start_letter, move_file_folder, move_folder_id, color='r')
    print_color(file_name, file_id, color='b')

    GdriveAPI.move_file(file_id=file_id, new_folder_id=move_folder_id)


    return True, move_file_folder


def get_new_files_to_send_out(x, GsheetAPI, GdriveAPI, child_folder_id,auto_publish_sheet_name ):
    print_color(child_folder_id, color='p')
    row_count = GsheetAPI.get_row_count(sheetname=auto_publish_sheet_name)
    GsheetAPI.sort_sheet( gid=0, sort_range= [0,1,17,row_count], dimensionIndex=11, sortOrder='ASCENDING')

    recruited_file_data = GsheetAPI.get_data_from_sheet(sheetname=auto_publish_sheet_name, range_name="A:L")
    lines_to_delete = recruited_file_data[(recruited_file_data['all_fields_assigned']=='FALSE') &
                                          (recruited_file_data['approved_to_send_out_?'] != 'TRUE')]
    lines_to_delete = lines_to_delete.iloc[::-1]

    lines_to_delete['index'] = lines_to_delete.index
    print_color(lines_to_delete, color='p')  # for i in range()
    print_color(lines_to_delete, color='r')

    for i in range(lines_to_delete.shape[0]):
        line_id = lines_to_delete['index'].iloc[i]
        print_color(line_id, color='r')
        GsheetAPI.delete_row_from_sheet(gid=0,delete_range= ['A',line_id,'Q',line_id])
        # break

    recruited_file_data = GsheetAPI.get_data_from_sheet(sheetname=auto_publish_sheet_name, range_name="A:L")
    if recruited_file_data.shape[0] >0:
        recruited_files = recruited_file_data['document_name'].unique()
    else:
        recruited_files = []
    print_color(recruited_files, color='y')
    print_color(recruited_file_data, color='r')

    if recruited_file_data.shape[0] == 0:
        max_id = 0
    else:
        recruited_file_data['id'] = recruited_file_data['id'].astype(int)
        max_id = int(recruited_file_data['id'].max())

    print_color(f'max_id: {max_id}', color='r')
    print_color(child_folder_id, color='y')

    files = GdriveAPI.get_files(folder_id=child_folder_id)
    print_color(files, color='y')
    print_color(len(files), color='y')

    all_pending_documents = [x for x in files]
    files_dict = {item['id']: item['name'] for item in all_pending_documents if ".doc" in item['name']}



    print_color(files_dict, color='b')
    for key, val in files_dict.items():
        print_color(key, val, color='r')
        max_id += 1
        process_doc_file(x=x, GdriveAPI=GdriveAPI, GsheetAPI=GsheetAPI, file_id=key, file_name=val, id_number=max_id)
        time.sleep(2)
        # break

    # return all_pending_documents


def email_approved_files(x, environment, GdriveAPI, GsheetAPI, GmailAPI, child_folders, auto_publish_sheet_name, child_folder_id):
    files = GdriveAPI.get_files(folder_id=child_folder_id)
    print_color(files, color='y')
    print_color(len(files), color='y')

    all_pending_documents = [x for x in files]
    print_color(all_pending_documents, color='y')

    file_data = GsheetAPI.get_data_from_sheet(sheetname=auto_publish_sheet_name, range_name="A:Q")
    print_color(file_data, color='g')
    pending_documents = [x['name'] for x in all_pending_documents]
    data_approved_to_email = file_data[(file_data['approved_to_send_out_?'] == 'TRUE')
                                       & (file_data['document_emailed'] != 'TRUE')
                                       & (file_data['document_name'].isin(pending_documents))
                                       & (file_data['attorney_email'].str.contains(".com"))

                                       ]
    data_approved_to_email = data_approved_to_email.iloc[::-1]

    print_color(data_approved_to_email, color='y')
    for i in range(data_approved_to_email.shape[0]):
        id  = data_approved_to_email['id'].iloc[i]
        file_id =  data_approved_to_email['file_id'].iloc[i]
        file_name = data_approved_to_email['document_name'].iloc[i]
        date_of_report = data_approved_to_email['date_of_report'].iloc[i]
        attorney_name = data_approved_to_email['attorney_name'].iloc[i]
        attorney_email = data_approved_to_email['attorney_email'].iloc[i]
        patient_name = data_approved_to_email['patient_name'].iloc[i]
        dob = data_approved_to_email['dob'].iloc[i]
        doa = data_approved_to_email['doa'].iloc[i]

        file_converted = False
        email_sent = False
        file_moved = False
        new_file_folder = None

        file_converted, pdf_export, file_type = convert_doc_to_pdf(x, file_name)
        subject = file_name.split(f'.{file_type}')[0]

        if environment == 'development':
            attorney_email = 'admin@Simpletowork.com'
        # elif environment == 'production':
        #     attorney_email = 'admin@Simpletowork.com'

        print_color(id, file_id, color='y')

        if file_converted is True:
            email_sent = email_doc_out(x, GmailAPI, subject, date_of_report, attorney_name,attorney_email, patient_name, dob, doa, pdf_export)
            print_color(email_sent, color='r')
            file_moved = False
            new_file_folder = None
            if email_sent is True:
                file_moved, new_file_folder = move_drive_file(GdriveAPI, file_name, file_id, child_folders)

        update_google_sheet_record(GsheetAPI, id, file_id, file_converted, email_sent, file_moved, new_file_folder)
        # break


def run_email_process(x, environment):
    sheet_id = x.google_sheet_published
    auto_publish_sheet_name = x.auto_publish_sheet_name

    GdriveAPI = GoogleDriveAPI(credentials_file=x.drive_credentials_file, token_file=x.drive_token_file, scopes=x.drive_scopes)
    GsheetAPI = GoogleSheetsAPI(credentials_file=x.gsheet_credentials_file, token_file=x.gsheet_token_file, scopes=x.gsheet_scopes,sheet_id=sheet_id)
    GmailAPI = GoogleGmailAPI(credentials_file=x.gmail_credentials_file, token_file=x.gmail_token_file, scopes=x.gmail_scopes)

    folder_name = x.published_folder
    folders = GdriveAPI.get_drive_folder(folder_name=folder_name)
    folder_id = folders[0].get("id")

    child_folders = GdriveAPI.get_child_folders(folder_id=folder_id)
    child_folder_id = None
    for each_folder in child_folders:
        if each_folder.get("name") == x.sub_published_folder:
            child_folder_id = each_folder.get("id")
            break

    ''' get new files that need to be send out
    check which files have alreay been recruited
    get the difference and import into google sheet
    '''
    get_new_files_to_send_out(x, GsheetAPI, GdriveAPI, child_folder_id, auto_publish_sheet_name)
    ''' email approved files
    check which pending files are approved to send out
    convert file to pdf
    email file out to attorney
    move file from auto publish folder to storage folder
    update google sheet accordingly    
    '''
    email_approved_files(x, environment, GdriveAPI, GsheetAPI, GmailAPI, child_folders, auto_publish_sheet_name, child_folder_id)



