import re
import os
from zipfile import ZipFile
from pypdf import PdfMerger
import PyPDF2
import win32com.client
import fitz
import datetime
from reportlab.pdfgen import canvas
from docx2pdf import convert
import pypandoc
import pandas as pd
from sqlalchemy import inspect
from global_modules import print_color, create_folder, run_sql_scripts, Get_SQL_Types, engine_setup
from google_drive_class import GoogleDriveAPI
from google_sheets_api import GoogleSheetsAPI
import zipfile
import pprint
import time
import shutil
from docx import Document


def database_setup(engine, database_name):
    scripts = []
    scripts.append(f'Create database if not exists {database_name}')
    run_sql_scripts(engine=engine, scripts=scripts)


def table_setup(engine):
    scripts = []
    scripts.append(f'''Create Table if not exists Folders(
        Folder_ID varchar(75),
        Folder_Name varchar(255),
        New_Files_Imported boolean,
        Zip_Files_Exists boolean,
        Zip_Files_Unzipped boolean,
        PDF_File_Processed boolean    
    )''')

    scripts.append(f'''Create Table if not exists Files(
        File_ID varchar(75),
        File_Name Text,
        Folder_ID Varchar(75),
        Folder_Name Varchar(255)    
    )''')

    scripts.append(f'''Create Table if not exists merge_process(
        ID int auto_increment unique,
        Import_Date date,
        Folder_Name varchar(255),
        Folder_ID varchar(65),
        Link_Folder varchar(255),
        Is_Single_File boolean,
        Has_Zip_Files boolean,
        Zip_File_Unpacked boolean,
        Index_Page_Created boolean,
        File_Combined boolean,
        File_Exists_in_Record_Input_Processed_Inputs boolean
    )''')




    run_sql_scripts(engine=engine, scripts=scripts)


def unzip_files(zip_path, unzip_path):
    with ZipFile(zip_path, 'r') as zObject:
        zObject.extractall(path=unzip_path)

    print_color(f'Zip Folder Extracted', color='g')


def number_list():
    numbers = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10',
                   '11', '12', '13', '14', '15', '16', '17', '18', '19', '20',
                   '21', '22', '23', '24', '25', '26', '27', '28', '29', '30',
                   '31', '32', '33', '34', '35', '36', '37', '38', '39', '40',
                   '41', '42', '43', '44', '45', '46', '47', '48', '49', '50',
                   '51', '52', '53', '54', '55', '56', '57', '58', '59', '60',
                   '61', '62', '63', '64', '65', '66', '67', '68', '69', '70',
                   '71', '72', '73', '74', '75', '76', '77', '78', '79', '80',
                   '81', '82', '83', '84', '85', '86', '87', '88', '89', '90',
                   '91', '92', '93', '94', '95', '96', '97', '98', '99', '100']
    return numbers


def get_pdf_page_count(file_path):
    with open(file_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        page_count = len(pdf_reader.pages)
        return page_count


def convert_doc_to_docx(doc_path, docx_path):
    # Create a new instance of Word application
    word = win32com.client.Dispatch('Word.Application')
    doc = word.Documents.Open(doc_path)
    doc.SaveAs(docx_path, FileFormat=16)  # 16 corresponds to .docx format
    doc.Close()
    word.Quit()

def get_docx_page_count(file_path):
    'https://pandoc.org/installing.html'
    # if file_extension == 'doc':
    #     new_file_path = f'{".doc".join(file_path.split(".doc")[:-1])}.docx'
    #
    # else:
    #     new_file_path = file_path
    #     # text_content = pypandoc.convert_file(file_path, 'plain', format='docx')
    #     #
    #     # # Count the pages (you may need to adjust this based on your document's structure)
    #     # page_count = text_content.count('\f') + 1
    #
    #     # result = subprocess.run(['antiword', file_path], stdout=subprocess.PIPE, text=True)
    #     # text_content = result.stdout
    #
    #     # Count the pages (you may need to adjust this based on your document's structure)
    #     # page_count = text_content.count('\f') + 1
    # # else:
    doc = Document(file_path)
    page_count = sum(1 for _ in doc.element.xpath('//w:sectPr'))
    return page_count


def get_existing_patient_folders(GdriveAPI=None, response_folder_id=None, include_processed_folders=False):
    sub_child_folders = GdriveAPI.get_child_folders(folder_id=response_folder_id)
    print_color(sub_child_folders, color='r')
    sub_child_folders = [x for x in sub_child_folders if x.get("trashed") is False]
    # print_color(sub_child_folders, color='g')

    processed_folder_id = None
    for each_folder in sub_child_folders:
        if each_folder.get("name") == 'Processed Inputs':
            processed_folder_id = each_folder.get("id")
            sub_child_folders.remove(each_folder)
            break

    existing_patient_folders = {}
    for each_folder in sub_child_folders:
        folder_name = each_folder.get('name').strip()
        folder_id = each_folder.get('id')
        if folder_name in existing_patient_folders.keys():
            existing_patient_folders[folder_name].append(folder_id)
        else:
            existing_patient_folders.update({folder_name: [each_folder.get('id')]})

    if include_processed_folders is True:
        processed_folders = GdriveAPI.get_child_folders(folder_id=processed_folder_id)
        for each_folder in processed_folders:
            folder_name = each_folder.get('name').strip()
            folder_id = each_folder.get('id')
            if folder_name in existing_patient_folders.keys():
                existing_patient_folders[folder_name].append(folder_id)
            else:
                existing_patient_folders.update({folder_name: [each_folder.get('id')]})

    return existing_patient_folders


def rename_existing_folders(GdriveAPI, existing_patient_folders):
    numbers = number_list()
    print_color(existing_patient_folders, color='y')
    print_color(len(existing_patient_folders.keys()), color='g')
    counter = 0
    for key, val in existing_patient_folders.items():
        core_folder_name = key.strip().replace("  ", " ")
        if core_folder_name[-2:].strip() in numbers:
            core_folder_name = core_folder_name[:-2].strip()

        core_folder_name = re.sub(r'(?<=[,])(?=[^\s])', r' ', core_folder_name)
        # core_folder_name = core_folder_name.replace(" ",", ").replace(",,","")
        print_color(core_folder_name, color='r')
        if key.strip() != core_folder_name:
            counter +=1
            # print_color(core_folder_name, color='p')
            for each_item in val:
                print_color(key, "            ", core_folder_name, each_item, color='y')
                GdriveAPI.rename_folder(folder_id=each_item, new_folder_name=core_folder_name)
    print_color(f'{counter} Folders Renamed', color='g')


def merge_existing_folders(GdriveAPI, existing_patient_folders, response_folder_id):
    existing_patient_folders_with_duplicates = {}
    for key, val in existing_patient_folders.items():
        if len(val) > 1:
            print_color(key, val, color='g')
            existing_patient_folders_with_duplicates.update({key: val})

    print_color(len(existing_patient_folders_with_duplicates.keys()), color='b')
    print_color(existing_patient_folders_with_duplicates.keys(), color='g')
    counter = 0
    for key, val in existing_patient_folders_with_duplicates.items():
        original_folder_id = ''
        for each_item in val:
            data = GdriveAPI.get_file_data(file_id=each_item)
            print_color(data.get("owners")[0].get("emailAddress"), data.get("parents")[0], color='r')

            if data.get("parents")[0] != response_folder_id:
                print_color(f'Moving Folder to Record Input', color='b')
                GdriveAPI.move_file(file_id=each_item, new_folder_id=response_folder_id)

            if data.get("owners")[0].get("emailAddress") != 'asnmedico@gmail.com' and data.get("parents")[0] != response_folder_id:
                original_folder_id = each_item

            elif data.get("owners")[0].get("emailAddress") != 'asnmedico@gmail.com':
                original_folder_id = each_item

        print_color(original_folder_id, color='g')

        print_color(f'Folder has Duplicate Entries. Will Merge', color='p')
        if original_folder_id == "":
            original_folder_id = val[0]
        folders_to_process = val
        folders_to_process.remove(original_folder_id)
        # print_color(folders_to_process, color='p')
        for each_item in folders_to_process:
            print_color(each_item, color='p')


            ''' Get Files in Folder / Move Files / Remove Folder '''
            print_color(each_item, original_folder_id, color='b')
            folder_files = GdriveAPI.get_files(each_item)
            for each_file in folder_files:
                each_file_id = each_file.get("id")

                print_color(each_file, color='y')
                GdriveAPI.move_file(file_id=each_file_id, new_folder_id=original_folder_id)

            GdriveAPI.delete_folder(folder_id=each_item, folder_name=key)
            counter += 1
        # break
    print_color(f'{counter} Folders Merged', color='g')


def import_new_folders(engine, database_name, existing_patient_folders):
    table_name = 'folders'
    if inspect(engine).has_table(table_name):
        data_df = pd.read_sql(f'Select "SQL" as `TYPE`, Folder_ID, Folder_Name from {table_name}', con=engine)
    else:
        data_df = pd.DataFrame()

    print_color(data_df, color='r')
    print_color(existing_patient_folders, color='g')

    df = pd.DataFrame.from_dict(existing_patient_folders, orient='columns').transpose()
    df['Folder_Name'] = df.index
    df.insert(0, "TYPE", "Google Drive")
    df = df.reset_index(drop=True)

    df.columns = ["TYPE", 'Folder_ID', 'Folder_Name']
    print_color(df, color='g')
    # df = df[[ 'Folder_ID', 'Folder_Name']]

    existing_folders_df = df

    df = pd.concat([df, data_df]).drop_duplicates(subset=['Folder_ID'], keep=False)
    df = df[df['TYPE'] == "Google Drive"]
    print_color(df, color='p')

    df = df.drop(columns=['TYPE'])

    sql_types = Get_SQL_Types(df).data_types
    df.to_sql(name=table_name, con=engine, if_exists='append', index=False, schema=database_name, chunksize=1000,
              dtype=sql_types)

    data_df = pd.read_sql(f'Select "SQL" as `TYPE`, Folder_ID, Folder_Name from {table_name}', con=engine)
    new_df = data_df.merge(existing_folders_df, left_on='Folder_ID', right_on='Folder_ID', how='left')
    print_color(new_df, color='y')
    for i in range(new_df.shape[0]):
        check_exists = new_df['TYPE_y'].iloc[i]
        folder_id = new_df['Folder_ID'].iloc[i]
        print_color(check_exists, color='r')
        if str(check_exists) == 'nan':
            run_sql_scripts(engine=engine, scripts=[f'Delete from folders where folder_id = "{folder_id}"'])


def process_new_files(engine, GdriveAPI, response_folder_id, existing_patient_folders):

    merge_process_df = pd.read_sql(f'Select * from merge_process', con=engine)

    numbers = number_list()
    all_files = GdriveAPI.get_files(folder_id=response_folder_id)
    print_color(len(all_files), color='y')
    all_files= [x for x in all_files if "Combined.pdf" not in x.get('name')]
    print_color(len(all_files), color='y')

    folder_dict = {}

    for i, each_file in enumerate(all_files):

        core_file_name = ".".join(each_file.get("name").split(".")[:-1])
        print_color(i, core_file_name, each_file.get("name"), color='g')
        file_id = each_file.get("id")
        file_extension = each_file.get("name").split(".")[-1]
        if core_file_name[-2:].strip() in numbers:
            core_file_name = core_file_name[:-2]
            print_color(i, core_file_name, color='r')
            core_file_name = re.sub(r'(?<=[,])(?=[^\s])', r' ', core_file_name)
            if core_file_name in folder_dict.keys():
                folder_dict[core_file_name]["ids"].append(file_id)
                folder_dict[core_file_name]["extensions"].append(file_extension)
            else:
                folder_dict.update({core_file_name: {"ids": [file_id], "extensions": [file_extension]}})
        else:
            core_file_name = re.sub(r'(?<=[,])(?=[^\s])', r' ', core_file_name)
            print_color(i, core_file_name, color='y')
            if core_file_name in folder_dict.keys():
                folder_dict[core_file_name]["ids"].append(file_id)
                folder_dict[core_file_name]["extensions"].append(file_extension)
            else:
                folder_dict.update({core_file_name:  {"ids": [file_id], "extensions": [file_extension]}})

    folder_dict = dict(sorted(folder_dict.items()))
    print_color(folder_dict, color='g')


    ''' CORE NAME = DATE, LAST NAME, FIRST NAME'''
    ''' ONLY MOVE FILES THAT HAVE MORE THAN 1 FILE PER CORE NAME'''
    patient_folders = list(existing_patient_folders.keys())
    for key, val in folder_dict.items():
        print_color(key, val, color='p')
        if key in patient_folders:
            print_color(f'Folder Already Exists', color='y')
            parent_folder_id = existing_patient_folders.get(key)[0]
            data = GdriveAPI.get_file_data(parent_folder_id)
            print_color(data.get("parents")[0] , response_folder_id, color='g')
            if data.get("parents")[0] != response_folder_id:
               GdriveAPI.move_file(file_id=parent_folder_id, new_folder_id=response_folder_id)
               scripts = [f'''insert into folders(`Folder_ID`, `Folder_Name`, `New_Files_Imported`)
                               values("{parent_folder_id}", "{key}", True)''']
               run_sql_scripts(engine=engine, scripts=scripts)

            for each_id in val.get("ids"):
                file_df = merge_process_df[(merge_process_df['Folder_ID'] == each_id)]
                if file_df.shape[0] > 0:
                    scripts.append(f'Delete from merge_process where Folder_ID = "{each_id}"')
                GdriveAPI.move_file(file_id=each_id, new_folder_id=parent_folder_id)
            scripts = [f'Update folders set New_Files_Imported = TRUE, PDF_File_Processed = null where Folder_ID="{parent_folder_id}"']
            run_sql_scripts(engine=engine, scripts=scripts)
        else:
            if len(val.get("ids")) > 1:
                print_color(f'Folder Does Not Exists', color='r')
                folder_id = GdriveAPI.create_folder(folder_name=key, parent_folder=response_folder_id)

                scripts = []
                for each_id in val.get("ids"):
                    file_df = merge_process_df[(merge_process_df['Folder_ID'] == each_id)]
                    if file_df.shape[0] >0:
                        scripts.append(f'Delete from merge_process where Folder_ID = "{each_id}"')
                    GdriveAPI.move_file(file_id=each_id, new_folder_id=folder_id)


                scripts.append(f'''insert into folders(`Folder_ID`, `Folder_Name`, `New_Files_Imported`)
                        values("{folder_id}", "{key}", True)''')
                if "zip" in val.get("extensions"):
                    has_zip = True
                else:
                    has_zip = False
                scripts.append(f'''insert into merge_process
                      values(null, curdate(), "{key}", "{folder_id}", "https://drive.google.com/file/d/{folder_id}", False, {has_zip}, False, False, False, False)
                                                                           ''')

                run_sql_scripts(engine=engine, scripts=scripts)
            else:
                print_color(f'File is Single File will not move', color='r')
                file_id = val.get("ids")[0]
                file_extension = val.get("extensions")[0]
                file_df = merge_process_df[(merge_process_df['Folder_ID'] == file_id)]
                if file_extension != "zip":
                    if file_df.shape[0] == 0:
                        scripts = [f'''insert into merge_process
                            values(null, curdate(), "{key}", "{file_id}", "https://drive.google.com/file/d/{file_id}", True, False, False, False, False, False)
                            ''']
                        run_sql_scripts(engine=engine, scripts=scripts)

        # break



    ''' PROCESS SINGLE ZIP FILES'''
    all_files = GdriveAPI.get_files(folder_id=response_folder_id)
    print_color(len(all_files), color='y')
    all_files = [x for x in all_files if "Combined.pdf" not in x.get('name')]
    for i, each_file in enumerate(all_files):
        print_color(i, each_file, color='g')
        core_file_name = ".".join(each_file.get("name").split(".")[:-1])
        file_id = each_file.get("id")
        file_extension = each_file.get("name").split(".")[-1].lower()
        if core_file_name[-2:].strip() in numbers:
            core_file_name = core_file_name[:-2]
            print_color(core_file_name, color='r')
            core_file_name = re.sub(r'(?<=[,])(?=[^\s])', r' ', core_file_name)
            if core_file_name in folder_dict.keys():
                folder_dict[core_file_name].append(file_id)
            else:
                folder_dict.update({core_file_name: [file_id]})
        else:
            core_file_name = re.sub(r'(?<=[,])(?=[^\s])', r' ', core_file_name)

        if file_extension == 'zip':
            if core_file_name in patient_folders:
                print_color(f'Folder Already Exists', color='y')
                new_folder_id = existing_patient_folders.get(core_file_name)[0]

                # for each_id in val:
                GdriveAPI.move_file(file_id=file_id, new_folder_id=new_folder_id)
                scripts = [f'Update folders set New_Files_Imported = TRUE, PDF_File_Processed = null where Folder_ID="{new_folder_id}"']
                run_sql_scripts(engine=engine, scripts=scripts)
            else:
                print_color(f'Folder Does Not Exists', color='r')
                folder_id = GdriveAPI.create_folder(folder_name=core_file_name, parent_folder=response_folder_id)

                GdriveAPI.move_file(file_id=file_id, new_folder_id=folder_id)
                scripts = [f'''insert into folders(`Folder_ID`, `Folder_Name`, `New_Files_Imported`)
                                      values("{folder_id}", "{core_file_name}", True)''']
                run_sql_scripts(engine=engine, scripts=scripts)
        else:
            if core_file_name in patient_folders:
                print_color(f'Folder Already Exists', color='y')
                new_folder_id = existing_patient_folders.get(core_file_name)[0]

                GdriveAPI.move_file(file_id=file_id, new_folder_id=new_folder_id)
                scripts = [f'Update folders set New_Files_Imported = TRUE, PDF_File_Processed = null  where Folder_ID="{new_folder_id}"']
                run_sql_scripts(engine=engine, scripts=scripts)


def map_files_and_folders_to_google_drive(GdriveAPI, GsheetAPI, response_folder_id):
    # sheet_name = "Merged Files - Record Input"
    # df = GsheetAPI.get_data_from_sheet(sheetname=sheet_name, range_name="A:K")

    sub_child_folders = GdriveAPI.get_child_folders(folder_id=response_folder_id)

    processed_folder_id = None
    for each_folder in sub_child_folders:
        if each_folder.get("name") == 'Processed Inputs':
            processed_folder_id = each_folder.get("id")
            sub_child_folders.remove(each_folder)
            break

    all_files = GdriveAPI.get_files(folder_id=response_folder_id)
    print_color(len(all_files), color='y')
    all_files= [x for x in all_files if "Combined.pdf" not in x.get('name')]

    # row_number = GsheetAPI.get_row_count(sheetname=sheet_name)

    single_file_data = []
    for each_file in all_files:
        print_color(each_file, color='g')
        single_file_dict = {
            "id": None,
            "import_date": datetime.datetime.now().strftime("%Y-%m-%d"),
            "folder_name":  each_file.get("name"),
            "folder_id": each_file.get("id"),
            "link_to_folder": f"https://drive.google.com/file/d/{each_file.get('id')}/",
            "is_single_file": True,
            "has_zip_files": False,
            "zip_file_unpacked": False,
            "index_page_created": False,
            "file_combined": False,
            "file_exists_in_record_input": False,


        }
        single_file_data.append(single_file_dict)
        break






def move_files_to_parent_folder(root_folder):
    # Iterate through all subdirectories
    for subdir, _, files in os.walk(root_folder):
        for file in files:
            # Construct the paths for the source and destination
            source_path = os.path.join(subdir, file)
            destination_path = os.path.join(root_folder, file)

            # Move the file to the parent folder
            shutil.move(source_path, destination_path)
            print(f"Moved: {source_path} to {destination_path}")


def unpack_child_folders(GdriveAPI, parent_folder, processed_folder_id, child_folders):
    for each_folder in child_folders:
        child_folder_id = each_folder.get("id")
        child_folder_name =  each_folder.get("name")
        sub_child_folders = GdriveAPI.get_child_folders(folder_id=child_folder_id)
        if len(sub_child_folders) >0:
            unpack_child_folders(GdriveAPI, parent_folder, processed_folder_id, sub_child_folders)
        child_folder_files = GdriveAPI.get_files(folder_id=child_folder_id)
        for each_child_file in child_folder_files:
            child_file_id = each_child_file.get("id")
            GdriveAPI.move_file( file_id=child_file_id, new_folder_id=parent_folder)

        try:
            GdriveAPI.delete_folder(file_id=child_folder_id, folder_name=child_folder_name)
        except:
            print_color(f'Could Not Delete Folder. WIll Move instad', color='r')
            GdriveAPI.move_file( file_id=child_folder_id, new_folder_id=processed_folder_id)


def process_zip_files(GdriveAPI, file_export, folder_id, processed_folder_id, file_id, file_name, extended_file_name, viewable_files):
    numbers = number_list()
    print_color(f'Run Zip Process', color='y')
    print_color(file_name, color='r')

    core_file_name = ".".join(file_name.split(".")[:-1])
    number_assignment = None
    # file_id = each_file.get("id")
    if core_file_name[-2:].strip() in numbers:
        number_assignment = core_file_name[-2:].strip()
        core_file_name = core_file_name[:-2]
        print_color(core_file_name, color='r')
        core_file_name = re.sub(r'(?<=[,])(?=[^\s])', r' ', core_file_name)

    else:
        core_file_name = re.sub(r'(?<=[,])(?=[^\s])', r' ', core_file_name)

    print_color(number_assignment, color='p')
    existing_files = [x.get("name") for x in viewable_files]
    print_color(existing_files, color='y')

    GdriveAPI.download_file(file_id=file_id, file_name=extended_file_name)

    extended_unzipped_folder = f'{file_export}\\{".".join(file_name.split(".")[:-1])} unzipped'
    create_folder(extended_unzipped_folder)

    with zipfile.ZipFile(extended_file_name, 'r') as z:
        # Replace 'extracted_folder_id' with the desired folder ID for the extracted contents
        z.extractall(extended_unzipped_folder)
        print_color(f'Files Unzipped into {extended_unzipped_folder}', color='b')

        move_files_to_parent_folder(extended_unzipped_folder)
        extracted_files = os.listdir(extended_unzipped_folder)
        print_color(extracted_files, color='y')
        for each_item in extracted_files:
            if os.path.isdir(f'{extended_unzipped_folder}\\{each_item}'):
                shutil.rmtree(f'{extended_unzipped_folder}\\{each_item}')

    extracted_files = os.listdir(extended_unzipped_folder)
    for each_unzipped_file in extracted_files:
        file_path = f'{extended_unzipped_folder}\\{each_unzipped_file}'
        extension = each_unzipped_file.split(".")[-1]
        if number_assignment is not None:
            new_file_name = f'{".".join(each_unzipped_file.split(".")[:-1])} {number_assignment}.{extension}'
        else:
            new_file_name = each_unzipped_file
        if new_file_name in existing_files:
            print_color(f'File Already Uploaded', color='r')
        else:

            GdriveAPI.upload_file(folder_id=folder_id, file_name=new_file_name, file_path=file_path)
    delete_success = GdriveAPI.delete_file(file_id=file_id, file_name=file_name)
    if delete_success is False:
        GdriveAPI.move_file(file_id=file_id, new_folder_id=processed_folder_id)

    time.sleep(3)
    os.remove(extended_file_name)
    # shutil.rmtree(extended_file_name)
    shutil.rmtree(extended_unzipped_folder)


def sort_files(folder_files):
    numbers = number_list()
    for each_file in folder_files:
        file_name = each_file.get("name")
        core_file_name = ".".join(file_name.split(".")[:-1])
        number_assignment = 0
        # file_id = each_file.get("id")
        if core_file_name[-2:].strip() in numbers:
            number_assignment = core_file_name[-2:].strip()
            core_file_name = core_file_name[:-2]
            print_color(core_file_name, color='r')
            core_file_name = re.sub(r'(?<=[,])(?=[^\s])', r' ', core_file_name)

        else:
            core_file_name = re.sub(r'(?<=[,])(?=[^\s])', r' ', core_file_name)

        each_file['number_assignment'] = int(number_assignment)
        each_file['core_file_name'] = core_file_name

        print_color(folder_files, color='g')

    sorted_files = sorted(folder_files, key=lambda x: (x['number_assignment'], x['created_time']))
    # for each_file in sorted_files:
    #     file_name = each_file.get("name")
    #     print_color(file_name, color='y')

    return sorted_files


def bytes_to_gb(bytes_value):
    gb_value = bytes_value / (1024 ** 3)
    return gb_value


def get_file_size(sorted_files, extension_list):
    numbers = number_list()
    print_color(sorted_files, color='b')
    data = {
        'ID': [],
        'File Name': [],
        'File Extension': [],
        'File Size': []
    }

    considered_file_df = pd.DataFrame(data)
    print_color(considered_file_df, color='b')

    for each_file in sorted_files:
        # print_color(each_file, color='r')
        file_name = each_file.get("name")
        file_size = each_file.get("size")
        file_id = each_file.get("id")
        file_ext = each_file.get("file_extension")

        new_row = {'ID': file_id, 'File Name': file_name, 'File Extension': file_ext, 'File Size': file_size}
        if file_ext in extension_list:
            check_if_exists = considered_file_df[(considered_file_df['File Name'] == file_name)
                               & (considered_file_df['File Extension'] == file_ext)
                               & (considered_file_df['File Size'] == file_size)]
            if check_if_exists.shape[0] ==0:
                considered_file_df = considered_file_df._append(new_row, ignore_index=True)

            print_color(file_name, file_size, file_id, file_ext, color='y')

    considered_file_df['File Size'] = considered_file_df['File Size'].astype(int)
    considered_file_df['File Size in GB'] = considered_file_df['File Size'].apply(lambda x: bytes_to_gb(x))

    print_color(considered_file_df, color='b')
    print_color(considered_file_df['File Size in GB'].dtype, color='b')

    combined_files_size = considered_file_df['File Size in GB'].sum()
    # combined_files_size = bytes_to_gb(files_size)

    print_color(combined_files_size, color='p')

    return combined_files_size


def create_index(output_path, file_list):
    index_content = "Index Page\n\n"

    for i, file_info in enumerate(file_list):
        file_name, start_page, page_count = file_info
        if start_page+2 == start_page +page_count +1:
            index_content += f"{i + 1}) {file_name}     -   Page {start_page + 2}\n"
        else:
            index_content += f"{i + 1}) {file_name}   -   Pages {start_page+2}-{start_page +page_count +1 }\n"

    print_color(index_content, color='y')

    pdf_canvas = canvas.Canvas(output_path, pagesize=(612, 792))
    pdf_canvas.setFont("Helvetica", 12)

    lines = index_content.split("\n")

    # Write text to the PDF with new lines
    y_position = 735
    for line in lines:
        pdf_canvas.drawString(75, y_position, line)
        y_position -= 15  # Adjust the vertical position for the next line


    # pdf_canvas.drawString(100, 750, index_content)

    # Save the PDF
    pdf_canvas.save()

    # pdf_writer.add_page(page=(0))
    # index_page = pdf_writer.getPage(pdf_writer.getNumPages() - 1)
    # index_page.mergePage(PyPDF2.PdfFileReader(index_content).getPage(0))


def merge_to_pdf(GdriveAPI, sorted_files, export_folder_name, folder_name, extension_list,
                       processed_folder_id, response_folder_id, folder_id):
    pdf_writer = PyPDF2.PdfWriter()
    combined_pdf_file = f'{export_folder_name}\\{folder_name} Combined Draft.pdf'
    if os.path.exists(combined_pdf_file):
        os.remove(combined_pdf_file)
    final_combined_pdf_file = f'{export_folder_name}\\{folder_name} Combined.pdf'
    if os.path.exists(final_combined_pdf_file):
        os.remove(final_combined_pdf_file)
    # merger = PdfMerger()
    merger = fitz.open()
    merged_files = []
    file_list = []
    start_page = 0

    for each_file in sorted_files:
        file_id = each_file.get("id")
        file_name = each_file.get("name")
        core_file_name = each_file.get("core_file_name")
        file_extension = each_file.get("file_extension")
        size = each_file.get("size")
        if file_extension in extension_list:
            print_color(size, color='y')
            export_file_name = f'{export_folder_name}\\{file_name}'
            print_color(export_file_name, color='y')
            if file_extension in ["jpg", "jpeg", "png"]:
                if int(size) < 1000:
                    print_color(f'Image is less than 1 KB so will ignore file', color='r')
                    return

            GdriveAPI.download_file(file_id=file_id, file_name=export_file_name)
            print_color(export_file_name, color='g')
            if file_extension.lower() == 'doc':
                new_file_path = f'{".doc".join(export_file_name.split(".doc")[:-1])}.docx'
                convert_doc_to_docx(doc_path=export_file_name, docx_path=new_file_path)
                export_file_name = new_file_path
                file_extension = 'docx'
            if file_extension.lower() in ['docx']:
                page_count = get_docx_page_count(export_file_name)
            elif file_extension.lower() in ['pdf']:
                page_count = get_pdf_page_count(export_file_name)
            file_details = f'{core_file_name}, {file_extension}, {size}, {page_count}'
            # print_color(file_details, color='y')
            # print_color(merged_files, color='y')
            if file_details in merged_files:
                print_color(f'File Already Exists and is a duplicate. Will not merge', color='r')
            else:
                merged_files.append(file_details)
                print_color(f'File Merged into Combined PDF', color='g')
                if file_extension != "pdf":
                    converted_export_file_name = f'{".".join(export_file_name.split(".")[:-1])}.pdf'
                    if file_extension in ['doc', 'docx']:
                        convert(export_file_name, converted_export_file_name)

                    adjusted_export_file_name = converted_export_file_name
                else:
                    adjusted_export_file_name = export_file_name

                pdf_document = fitz.open(adjusted_export_file_name)
                merger.insert_pdf(pdf_document)

                # merger.append(adjusted_export_file_name)

                file_list.append((file_name, start_page, page_count))
                start_page += page_count

            GdriveAPI.move_file(file_id=file_id, new_folder_id=processed_folder_id)

    index_path = f'{export_folder_name}\\index.pdf'

    create_index(index_path, file_list)


    print_color(merger, color='y')

    # merger.write(combined_pdf_file)
    merger.save(combined_pdf_file)
    merger.close()

    new_merger = PdfMerger()
    new_merger.append(index_path)
    new_merger.append(combined_pdf_file)
    new_merger.write(final_combined_pdf_file)
    new_merger.close()

    print_color('PDF Merged', color='y')
    upload_folder_id = response_folder_id
    # upload_folder_id = folder_id

    '''check if a combined file already exists. if so, delete'''
    final_upload_file_name = f'{folder_name} Combined.pdf'
    all_files = GdriveAPI.get_files(folder_id=response_folder_id)
    all_files = [x for x in all_files if x.get("name") == final_upload_file_name]
    if len(all_files) >0:
        GdriveAPI.delete_file(file_id=all_files[0].get("id"), file_name=final_upload_file_name)
    GdriveAPI.upload_file(folder_id=upload_folder_id, file_name=final_upload_file_name,
                          file_path=final_combined_pdf_file)





def process_open_folders(x, engine, GdriveAPI, response_folder_id):
    numbers = number_list()
    file_export = f'{x.project_folder}\\Record Inputs'
    create_folder(file_export)
    extension_list = ["pdf", "docx", "doc", "png", "jpeg", "jpg"]
    df = pd.read_sql(f'''Select * from folders where (New_Files_Imported is null or New_Files_Imported = 1)
        and (PDF_File_Processed != 1 or PDF_File_Processed is null)
        and Folder_Name not in ("1 - Folders For Review With Alan", "Processed Inputs", "Doubt Files", "Old Reports")
        and Folder_Name in ("2023.12.14, Dalbo, James")
        order by Folder_Name
    ''', con=engine)
    merge_process_df = pd.read_sql(f'Select * from merge_process', con=engine)


    print_color(df, color='r')

    for i in range(df.shape[0]):
        folder_id = df['Folder_ID'].iloc[i]
        folder_name = df['Folder_Name'].iloc[i].strip()
        new_files_imported = df['New_Files_Imported'].iloc[i]
        zip_files_exists = df['Zip_Files_Exists'].iloc[i]
        zip_files_unzipped = df['Zip_Files_Unzipped'].iloc[i]
        pdf_file_processed = df['PDF_File_Processed'].iloc[i]



        export_folder_name =  f'{file_export}\\{folder_name}'
        create_folder(export_folder_name)

        print_color(f'{i}/{df.shape[0]} Getting files for {folder_id}: {folder_name}')
        folder_files = GdriveAPI.get_files(folder_id)
        get_folder_sub_folders = GdriveAPI.get_child_folders(folder_id=folder_id)
        folder_files = [x for x in folder_files if x.get("trashed") == False]
        zip_files = [x for x in folder_files if "zip" in x.get("name").lower()]
        viewable_files = [x for x in folder_files if x.get("name").split(".")[-1].lower() in extension_list]

        has_zip = True if len(zip_files) >0 else False

        folder_ids = merge_process_df[(merge_process_df['Folder_ID'] == folder_id)]
        if folder_ids.shape[0] >0:
            scripts = [f'''update merge_process set  
                Import_Date = curdate(),
                Is_Single_File= False, 
                Has_Zip_Files = {has_zip},
                Zip_File_Unpacked = False, 
                Index_Page_Created = False, 
                File_Combined = False, 
                File_Exists_in_Record_Input_Processed_Inputs = False
                where Folder_ID ="{folder_id}"
                ''']
        else:
            scripts = [f'''insert into merge_process
                      values(null, curdate(), "{folder_name}", "{folder_id}", "https://drive.google.com/file/d/{folder_id}", False, {has_zip}, False, False, False, False)
                                                                       ''']
        run_sql_scripts(engine=engine, scripts=scripts)


        print_color(folder_files, color='y')
        print(f'Folder Count {len(folder_files)}' , f'Processed Folder Count {len(get_folder_sub_folders)}')
        if len(folder_files) == 0 and len(get_folder_sub_folders) == 0:
            GdriveAPI.delete_folder(folder_id=folder_id, folder_name=folder_name)
            scripts = []
            scripts.append(f'Delete from folders where Folder_ID = "{folder_id}"')
            scripts.append(f'Delete from merge_process where Folder_ID = "{folder_id}"')

            run_sql_scripts(engine=engine, scripts=scripts)

        elif len(folder_files) == 1 and len(get_folder_sub_folders) == 0 and len(zip_files) ==0:
            '''MOVE FILE OUT AS SINGLE FILE / REMOVE FOLDER'''
            ''' RENAME FILE TO CORE LOGIC'''
            each_file =  folder_files[0]

            # core_file_name = ".".join(folder_name.split(".")[:-1])
            file_id = each_file.get("id")
            file_extension = each_file.get("name").split(".")[-1]
            # if core_file_name[-2:].strip() in numbers:
            #     core_file_name = core_file_name[:-2]
            #     print_color(i, core_file_name, color='r')
            #     core_file_name = re.sub(r'(?<=[,])(?=[^\s])', r' ', core_file_name)
            # else:
            #     core_file_name = re.sub(r'(?<=[,])(?=[^\s])', r' ', core_file_name)
            print_color(folder_name, color='g')
            new_file_name = f'{folder_name}.{file_extension}'
            GdriveAPI.rename_file( file_id=file_id, new_file_name=new_file_name)
            GdriveAPI.move_file(file_id=file_id, new_folder_id=response_folder_id)
            GdriveAPI.delete_folder(folder_id=folder_id, folder_name=folder_name)
            scripts = []
            scripts.append(f'Delete from folders where Folder_ID = "{folder_id}"')
            scripts.append(f'Delete from merge_process where Folder_ID = "{folder_id}"')
            run_sql_scripts(engine=engine, scripts=scripts)

        else:
            print_color(folder_files, color='y')
            '''Step 1 - Check if there are files that we already processed'''
            '''Step 2 - Check if there are folders in the folder. If so unpack files into main folder'''

            '''Step 3 - Unzip any Zip Files *'''
            '''Step 4 - For Files Already Process get file content'''
            '''Step 5 - Get Files from Original Folder with New Zipped Files'''
            '''Step 6 - Combine Processed and Unprocessed Files'''
            '''Step 7 - Sort File By Create Date'''
            '''Step 8 - Get Combined File Size'''
            '''Step 9 - Merge Files To one PDF
                      - Create Index Page'''
            '''       - Move Files to "Processed Folder"'''

            '''Step 1 - Check if there are files that we already processed'''
            processed_folder_id = None
            child_folders = GdriveAPI.get_child_folders(folder_id=folder_id)
            for each_folder in child_folders:
                if "Processed Files" in each_folder.get("name"):
                    processed_folder_id = each_folder.get("id")
                    break

            if processed_folder_id is None:
                processed_folder_id = GdriveAPI.create_folder(folder_name='Processed Files', parent_folder=folder_id)


            '''Step 2 - Check if there are folders in the folder. If so unpack files into main folder'''
            additional_child_folders = [x for x in child_folders if x.get("name") != 'Processed Files']
            print_color(additional_child_folders, color='y')
            unpack_child_folders(GdriveAPI=GdriveAPI, parent_folder=folder_id, processed_folder_id=processed_folder_id,
                                 child_folders=additional_child_folders)

            '''Step 3 - Unzip any Zip Files *'''
            for each_file in folder_files:
                file_extension = each_file.get("file_extension")
                file_id = each_file.get("id")
                # pprint.pprint(GdriveAPI.get_file_data(file_id))
                file_name = each_file.get("name")
                extended_file_name = f'{file_export}\\{file_name}'
                print_color(each_file, color='g')
                if file_extension == 'zip':
                    process_zip_files(GdriveAPI, file_export, folder_id, processed_folder_id, file_id, file_name, extended_file_name,
                                      viewable_files)
                    scripts = [f'''update merge_process set Zip_File_Unpacked = True where Folder_ID ="{folder_id}"
                                   ''']
                    run_sql_scripts(engine=engine, scripts=scripts)
                    # break

            '''Step 4 - For Files Already Process get file content'''
            processed_folder_files = GdriveAPI.get_files(processed_folder_id)
            processed_folder_files = [x for x in processed_folder_files if x.get("trashed") == False]

            '''Step 5 - Get Files from Original Folder with New Zipped Files'''
            folder_files = GdriveAPI.get_files(folder_id)
            folder_files = [x for x in folder_files if x.get("trashed") == False]
            print_color(folder_files, color='b')
            if len(folder_files) ==0 and len(processed_folder_files) ==0 :
                print_color(f'No New Files to Process', color='r')
                break
            '''Step 6 - Combine Processed and Unprocessed Files'''
            folder_files = folder_files + processed_folder_files
            folder_files = [x for x in folder_files if x.get("trashed") == False]
            folder_files = [x for x in folder_files if x.get("file_extension") in extension_list]
            print_color(len(folder_files), color='y')

            '''Step 7 - Sort File By File Number - Create Date'''
            sorted_files = sort_files(folder_files)

            '''Step 8 - Get combined size of all unique files'''
            combined_files_size = get_file_size(sorted_files, extension_list)

            # '''Step 9 - Merge Files To one PDF'''
            # # print_color(len(sorted_files),color='y')
            # if combined_files_size > .80:
            #     print_color(f'Combined File Size in folder exceed Allowed Size to run', color='r')
            # else:
            #     merge_to_pdf(GdriveAPI, sorted_files, export_folder_name, folder_name, extension_list, processed_folder_id,
            #                  response_folder_id, folder_id)
            #
            #     scripts = []
            #     scripts.append(f'Update folders set PDF_File_Processed = True, New_Files_Imported=null where Folder_ID = "{folder_id}"')
            #
            #     run_sql_scripts(engine=engine, scripts=scripts)

        break


def merge_files_to_pdf(x, environment):
    engine = engine_setup(hostname=x.hostname, username=x.username, password=x.password, port=x.port)
    engine_1 = engine_setup(project_name=x.project_name, hostname=x.hostname, username=x.username, password=x.password,
                            port=x.port)
    database_name =x.project_name
    database_setup(engine=engine, database_name=database_name)
    table_setup(engine_1)

    GdriveAPI = GoogleDriveAPI(credentials_file=x.drive_credentials_file, token_file=x.drive_token_file,
                               scopes=x.drive_scopes)
    GsheetAPI = GoogleSheetsAPI(credentials_file=x.gsheet_credentials_file, token_file=x.gsheet_token_file,
                               scopes=x.gsheet_scopes, sheet_id=x.google_sheet_merge_process)

    folder_id = x.mle_folder
    print_color(folder_id, color='b')
    child_folders = GdriveAPI.get_child_folders(folder_id=folder_id)
    for each_folder in child_folders:
        if each_folder.get("name") == 'RECORD-INPUT':
            response_folder_id = each_folder.get("id")
            break

    print_color(response_folder_id, color='y')

    # '''GET DICT OF ALL FOLDERS IN RECORD-INPUT'''
    # existing_patient_folders = get_existing_patient_folders(GdriveAPI, response_folder_id)
    # print_color(len(existing_patient_folders), color='y')
    # '''RENAME FOLDERS THAT ARE NOT FORMATTED PROPERLY'''
    # rename_existing_folders(GdriveAPI, existing_patient_folders)
    # '''GET UPDATED DICT OF ALL FOLDERS IN RECORD-INPUT'''
    # existing_patient_folders = get_existing_patient_folders(GdriveAPI, response_folder_id, include_processed_folders=True)
    # print_color(len(existing_patient_folders), color='y')
    # '''MERGE FOLDERS THAT HAVE THE SAME NAME'''
    # merge_existing_folders(GdriveAPI, existing_patient_folders, response_folder_id)
    # '''GET UPDATED DICT OF ALL FOLDERS IN RECORD-INPUT'''
    # existing_patient_folders = get_existing_patient_folders(GdriveAPI, response_folder_id)
    # '''MAP FOLDERS TO SQL'''
    # import_new_folders(engine_1, database_name, existing_patient_folders)
    # '''GET UPDATED DICT OF ALL FOLDERS IN RECORD-INPUT'''
    # existing_patient_folders = get_existing_patient_folders(GdriveAPI, response_folder_id,include_processed_folders=True)
    # '''MAP / MOVE NEW FILES IN RECORD INPUT'''
    # process_new_files(engine_1, GdriveAPI, response_folder_id, existing_patient_folders)
    # '''MAP MERGE PROCESS TO GOOGLE SHEETS'''
    # map_files_and_folders_to_google_drive(GdriveAPI, GsheetAPI, response_folder_id)


    '''PROCESS FOLDERS THAT NEED TO BE MERGED TO A PDF'''
    process_open_folders(x, engine_1, GdriveAPI, response_folder_id)









