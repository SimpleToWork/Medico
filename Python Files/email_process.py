from google_drive_class import GoogleDriveAPI
from global_modules import print_color, create_folder
from google_sheets_api import GoogleSheetsAPI
from docx import Document
import docx2txt
from dateutil.parser import parse
import io
import getpass
import re

def is_date(string, fuzzy=False):
    """
    Return whether the string can be interpreted as a date.

    :param string: str, string to check for date
    :param fuzzy: bool, ignore unknown tokens in string if True
    """
    try:
        parse(string, fuzzy=fuzzy)
        return True

    except ValueError:
        return False

def upload_data_to_gmail(x):
    credentials_file = x.gmail_credentials_file
    token_file = x.gmail_token_file
    scopes = x.gmail_scopes
    sheet_id = x.google_sheet_published

    GsheetAPI = GoogleSheetsAPI(credentials_file=None, token_file=None, scopes=None, sheet_id=sheet_id)

    # df = GsheetAPI.get_data_from_sheet(sheetname='KYF ACHWORKS', range_name='A:E')
    # print_color(df, color='r')
    #
    # row_number = df.shape[0] + 2
    #
    # GsheetAPI.write_data_to_sheet(
    #     data=df,
    #       sheetname='KYF ACHWORKS',
    #       row_number=row_number,
    #       include_headers=False,
    #       clear_data=False)

def process_doc_file(x, GdriveAPI, file_id, file_name):
    create_folder(f'{x.project_folder}\\Extracts')
    file_export = f'{x.project_folder}\\Extracts\\{file_name}'

    # GdriveAPI.download_file(file_id=file_id, file_name=file_export)

    data_dict = {
        "attorney_name": None,
        "patient_name": None,
        "date_of_report": None,
        "attorney_email": None,
        "date_of_birth": None,
        "date_of_arrival": None
    }

    first_line_text = 0
    date_of_report = None
    attorney_name = None


    # doc = Document(file_export)
    text = docx2txt.process(file_export)
    text_list = text.split("\n")
    # print_color(text_list, color='y')


    for i, row_text in enumerate(text_list):
        if i <= 30:
            text = row_text
            print_color(text, color='y')


            if first_line_text ==0:
                if text.strip() != "":
                    first_line_text += 1


            # lines.append(paragraph.text)
            # print(paragraph.text)
            if first_line_text != 0 and date_of_report is None:
                date_of_report = text

                data_dict["date_of_report"] = date_of_report
            if "RE:" in text.upper():
                data_dict["patient_name"] = text.replace("RE:","").replace(",","").replace("\t","").strip()
            if "DOB" in text.upper():
                data_dict["date_of_birth"] =text.replace("DOB:","").replace(",","").replace("\t","").strip()
            if "DOA" in text.upper():
                data_dict["date_of_arrival"] = text.replace("DOA","").replace(",","").replace("\t","").strip()
            if "ESQ." in text.upper():
                attorney_name = text.upper()
                data_dict["attorney_name"] = attorney_name.replace("ESQ.","").replace(",","").strip()
            if "@" and ".com" in text:
                data_dict["attorney_email"] = text


    print_color(data_dict, color='r')




    # Iterate through the paragraphs in the document




    # Print the first 15 lines
    # for line in lines:
    #     print(line)

    # print_color(file)
    # doc = Document()
    #

    # print_color(x.project_folder, color='g')
    # if file is not None:
    #     file.getvalue().decode('utf-8')
    #     doc.add_paragraph(text_data)
    #     doc.save(f)
    # pass

def run_email_process(x):
    credentials_file = x.drive_credentials_file
    token_file = x.drive_token_file
    scopes = x.drive_scopes

    GdriveAPI = GoogleDriveAPI(credentials_file=credentials_file, token_file=token_file, scopes=scopes)

    folder_name = "Published Reports"
    folders = GdriveAPI.get_drive_folder(folder_name=folder_name)
    folder_id = folders[0].get("id")

    child_folders = GdriveAPI.get_child_folders(folder_id=folder_id)

    print_color(child_folders, color='p')

    for each_folder in child_folders:
        if each_folder.get("name") == 'Testing STW':
            child_folder_id = each_folder.get("id")
            files = GdriveAPI.get_files(folder_id=child_folder_id)
            print_color(files, color='y')

            files_dict = {item['id']: item['name'] for item in files if ".doc" in item['name']}

            print_color(files_dict, color='b')
            for key, val in  files_dict.items():
                print_color(key, val, color='r')
                process_doc_file(x=x, GdriveAPI=GdriveAPI, file_id= key, file_name = val)
                break
            # process_doc_file(x)