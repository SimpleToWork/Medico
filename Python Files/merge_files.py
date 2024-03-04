import re
import os
from zipfile import ZipFile
from pypdf import PdfMerger
import PyPDF2
import win32com.client
import fitz
import subprocess
import datetime
from docx2pdf import convert
import pandas as pd
from sqlalchemy import inspect
from global_modules import print_color, create_folder, run_sql_scripts, Get_SQL_Types, engine_setup, error_handler
from google_drive_class import GoogleDriveAPI
from google_sheets_api import GoogleSheetsAPI
import zipfile
import time
import shutil
import pprint
from docx import Document
from pyhtml2pdf import converter
from PIL import Image


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
        Folder_Moved_To_Processed_Inputs boolean,
        Folder_To_Large_To_Combine boolean
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



def convert_image_to_pdf(image_path, output_pdf_path):
    # Open the image file
    img = Image.open(image_path)

    # Convert the image to RGB if it's not already in that mode
    if img.mode != 'RGB':
        img = img.convert('RGB')

    # Save the image as PDF
    img.save(output_pdf_path, 'PDF', resolution=100.0)


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


def get_existing_patient_folders(GdriveAPI=None, record_input_folder_id=None, sub_child_folders=None,
                                 processed_folder_id=None, include_processed_folders=False, main_log_file=None):
    print_color('''GET DICT OF ALL FOLDERS IN RECORD-INPUT''', color='k', output_file=main_log_file)

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


def rename_existing_folders(GdriveAPI, existing_patient_folders, main_log_file):
    print_color('''RENAME FOLDERS THAT ARE NOT FORMATTED PROPERLY''', color='k', output_file=main_log_file)

    numbers = number_list()
    print_color(f'Count of Patient Folders {len(existing_patient_folders.keys())}', color='g', output_file=main_log_file)
    print_color(existing_patient_folders, color='y', output_file=main_log_file)
    counter = 0
    for key, val in existing_patient_folders.items():
        core_folder_name = key.strip().replace("  ", " ")
        if core_folder_name[-2:].strip() in numbers:
            core_folder_name = core_folder_name[:-2].strip()

        core_folder_name = re.sub(r'(?<=[,])(?=[^\s])', r' ', core_folder_name)
        # core_folder_name = core_folder_name.replace(" ",", ").replace(",,","")
        print_color(core_folder_name, color='r', output_file=main_log_file)
        if key.strip() != core_folder_name:
            counter +=1
            # print_color(core_folder_name, color='p')
            for each_item in val:
                print_color(key, "            ", core_folder_name, each_item, color='y', output_file=main_log_file)
                GdriveAPI.rename_folder(folder_id=each_item, new_folder_name=core_folder_name)
    print_color(f'{counter} Folders Renamed', color='g', output_file=main_log_file)


def merge_existing_folders(GdriveAPI, existing_patient_folders, record_input_folder_id, main_log_file):
    print_color( '''MERGE FOLDERS THAT HAVE THE SAME NAME''', color='k', output_file=main_log_file)

    existing_patient_folders_with_duplicates = {}
    for key, val in existing_patient_folders.items():
        if len(val) > 1:
            print_color(key, val, color='g', output_file=main_log_file)
            existing_patient_folders_with_duplicates.update({key: val})

    print_color(len(existing_patient_folders_with_duplicates.keys()), color='b', output_file=main_log_file)
    print_color(existing_patient_folders_with_duplicates.keys(), color='g', output_file=main_log_file)
    counter = 0
    for key, val in existing_patient_folders_with_duplicates.items():
        original_folder_id = ''
        for each_item in val:
            data = GdriveAPI.get_file_data(file_id=each_item)
            print_color(data.get("owners")[0].get("emailAddress"), data.get("parents")[0], color='r', output_file=main_log_file)

            if data.get("parents")[0] != record_input_folder_id:
                print_color(f'Moving Folder to Record Input', color='b', output_file=main_log_file)
                GdriveAPI.move_file(file_id=each_item, new_folder_id=record_input_folder_id)

            if data.get("owners")[0].get("emailAddress") != 'asnmedico@gmail.com' and data.get("parents")[0] != record_input_folder_id:
                original_folder_id = each_item

            elif data.get("owners")[0].get("emailAddress") != 'asnmedico@gmail.com':
                original_folder_id = each_item

        print_color(original_folder_id, color='g', output_file=main_log_file)

        print_color(f'Folder has Duplicate Entries. Will Merge', color='p', output_file=main_log_file)
        if original_folder_id == "":
            original_folder_id = val[0]
        folders_to_process = val
        folders_to_process.remove(original_folder_id)
        # print_color(folders_to_process, color='p')
        for each_item in folders_to_process:
            print_color(each_item, color='p', output_file=main_log_file)
            ''' Get Files in Folder / Move Files / Remove Folder '''
            print_color(f'Merging {each_item} into {original_folder_id}', color='b', output_file=main_log_file)
            folder_files = GdriveAPI.get_files(each_item)
            for each_file in folder_files:
                each_file_id = each_file.get("id")
                print_color(each_file, color='y', output_file=main_log_file)
                GdriveAPI.move_file(file_id=each_file_id, new_folder_id=original_folder_id)
            folder_files = GdriveAPI.get_files(each_item)
            if len(folder_files) >0:
                print_color(f'Files Still Exists in Folder to Merge', color='r', output_file=main_log_file)
            else:
                GdriveAPI.delete_folder(folder_id=each_item, folder_name=key)
            counter += 1


    print_color(f'{counter} Folders Merged', color='g', output_file=main_log_file)


def import_new_folders(engine, database_name, existing_patient_folders, main_log_file):
    print_color('''MAP FOLDERS TO SQL''', color='k', output_file=main_log_file)

    table_name = 'folders'
    if inspect(engine).has_table(table_name):
        data_df = pd.read_sql(f'Select "SQL" as `TYPE`, Folder_ID, Folder_Name from {table_name}', con=engine)
    else:
        data_df = pd.DataFrame()

    print_color(data_df, color='r', output_file=main_log_file)
    print_color(existing_patient_folders, color='g', output_file=main_log_file)

    df = pd.DataFrame.from_dict(existing_patient_folders, orient='columns').transpose()
    df['Folder_Name'] = df.index
    df.insert(0, "TYPE", "Google Drive")
    df = df.reset_index(drop=True)

    df.columns = ["TYPE", 'Folder_ID', 'Folder_Name']
    print_color(df, color='g', output_file=main_log_file)
    # df = df[[ 'Folder_ID', 'Folder_Name']]

    existing_folders_df = df

    df = pd.concat([df, data_df]).drop_duplicates(subset=['Folder_ID'], keep=False)
    df = df[df['TYPE'] == "Google Drive"]
    print_color(df, color='p', output_file=main_log_file)

    df = df.drop(columns=['TYPE'])

    sql_types = Get_SQL_Types(df).data_types
    df.to_sql(name=table_name, con=engine, if_exists='append', index=False, schema=database_name, chunksize=1000,
              dtype=sql_types)

    data_df = pd.read_sql(f'Select "SQL" as `TYPE`, Folder_ID, Folder_Name from {table_name}', con=engine)
    new_df = data_df.merge(existing_folders_df, left_on='Folder_ID', right_on='Folder_ID', how='left')
    print_color(new_df, color='y', output_file=main_log_file)
    for i in range(new_df.shape[0]):
        check_exists = new_df['TYPE_y'].iloc[i]
        folder_id = new_df['Folder_ID'].iloc[i]
        print_color(check_exists, color='r', output_file=main_log_file)
        if str(check_exists) == 'nan':
            run_sql_scripts(engine=engine, scripts=[f'Delete from folders where folder_id = "{folder_id}"'])


def process_new_files(engine, GdriveAPI, record_input_folder_id, existing_patient_folders, main_log_file):
    print_color('''MAP FOLDERS TO SQL''', color='k', output_file=main_log_file)
    merge_process_df = pd.read_sql(f'Select * from merge_process', con=engine)

    numbers = number_list()
    all_files = GdriveAPI.get_files(folder_id=record_input_folder_id)
    print_color(len(all_files), color='y', output_file=main_log_file)
    all_files= [x for x in all_files if "Combined" not in x.get('name')]
    print_color(len(all_files), color='y', output_file=main_log_file)
    print_color(all_files, color='y', output_file=main_log_file)


    folder_dict = {}

    for i, each_file in enumerate(all_files):
        core_file_name = ".".join(each_file.get("name").split(".")[:-1])
        print_color(i, core_file_name, each_file.get("name"), color='g', output_file=main_log_file)
        file_id = each_file.get("id")
        file_extension = each_file.get("name").split(".")[-1]
        if core_file_name[-2:].strip() in numbers:
            core_file_name = core_file_name[:-2]
            print_color(i, core_file_name, color='r', output_file=main_log_file)
            core_file_name = re.sub(r'(?<=[,])(?=[^\s])', r' ', core_file_name)
            if core_file_name in folder_dict.keys():
                folder_dict[core_file_name]["ids"].append(file_id)
                folder_dict[core_file_name]["extensions"].append(file_extension)
            else:
                folder_dict.update({core_file_name: {"ids": [file_id], "extensions": [file_extension]}})
        else:
            core_file_name = re.sub(r'(?<=[,])(?=[^\s])', r' ', core_file_name)
            print_color(i, core_file_name, color='y', output_file=main_log_file)
            if core_file_name in folder_dict.keys():
                folder_dict[core_file_name]["ids"].append(file_id)
                folder_dict[core_file_name]["extensions"].append(file_extension)
            else:
                folder_dict.update({core_file_name:  {"ids": [file_id], "extensions": [file_extension]}})

    folder_dict = dict(sorted(folder_dict.items()))
    print_color(folder_dict, color='g', output_file=main_log_file)


    ''' CORE NAME = DATE, LAST NAME, FIRST NAME'''
    ''' ONLY MOVE FILES THAT HAVE MORE THAN 1 FILE PER CORE NAME'''
    patient_folders = list(existing_patient_folders.keys())
    for key, val in folder_dict.items():
        print_color(key, val, color='p', output_file=main_log_file)
        if key in patient_folders:
            scripts = []
            print_color(f'Folder Already Exists', color='y', output_file=main_log_file)
            parent_folder_id = existing_patient_folders.get(key)[0]
            data = GdriveAPI.get_file_data(parent_folder_id)
            print_color(data.get("parents")[0] , record_input_folder_id, color='g', output_file=main_log_file)
            if data.get("parents")[0] != record_input_folder_id:
               GdriveAPI.move_file(file_id=parent_folder_id, new_folder_id=record_input_folder_id)
               scripts.append(f'''insert into folders(`Folder_ID`, `Folder_Name`, `New_Files_Imported`)
                               values("{parent_folder_id}", "{key}", True)''')
               run_sql_scripts(engine=engine, scripts=scripts)

            for each_id in val.get("ids"):
                file_df = merge_process_df[(merge_process_df['Folder_ID'] == each_id)]
                if file_df.shape[0] > 0:
                    scripts.append(f'Delete from merge_process where Folder_ID = "{each_id}"')
                GdriveAPI.move_file(file_id=each_id, new_folder_id=parent_folder_id)
            scripts.append(f'Update folders set New_Files_Imported = TRUE, PDF_File_Processed = null where Folder_ID="{parent_folder_id}"')
            run_sql_scripts(engine=engine, scripts=scripts)
        else:
            if len(val.get("ids")) > 1:
                print_color(f'Folder Does Not Exists', color='r', output_file=main_log_file)
                folder_id = GdriveAPI.create_folder(folder_name=key, parent_folder=record_input_folder_id)

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
                      values(null, curdate(), "{key}", "{folder_id}", "https://drive.google.com/folder/d/{folder_id}", False, {has_zip}, False, False, False, False, False)
                                                                           ''')

                run_sql_scripts(engine=engine, scripts=scripts)
            else:
                print_color(f'File is Single File will not move', color='r', output_file=main_log_file)
                file_id = val.get("ids")[0]
                file_extension = val.get("extensions")[0]
                file_df = merge_process_df[(merge_process_df['Folder_ID'] == file_id)]
                if file_extension != "zip":
                    if file_df.shape[0] == 0:
                        scripts = [f'''insert into merge_process
                            values(null, curdate(), "{key}", "{file_id}", "https://drive.google.com/file/d/{file_id}", True, False, False, False, False, False, False)
                            ''']
                        run_sql_scripts(engine=engine, scripts=scripts)

    ''' PROCESS SINGLE ZIP FILES'''
    print_color(folder_dict, color='y', output_file=main_log_file)
    pprint.pprint(folder_dict)

    all_files = GdriveAPI.get_files(folder_id=record_input_folder_id)
    print_color(len(all_files), color='y', output_file=main_log_file)
    all_files = [x for x in all_files if "Combined" not in x.get('name')]
    for i, each_file in enumerate(all_files):
        print_color(i, each_file, color='g', output_file=main_log_file)
        core_file_name = ".".join(each_file.get("name").split(".")[:-1])
        file_id = each_file.get("id")
        file_extension = each_file.get("name").split(".")[-1].lower()
        if core_file_name[-2:].strip() in numbers:
            core_file_name = core_file_name[:-2]
            print_color(core_file_name, color='r', output_file=main_log_file)
            core_file_name = re.sub(r'(?<=[,])(?=[^\s])', r' ', core_file_name)
            if core_file_name in folder_dict.keys():
                folder_dict[core_file_name]["ids"].append(file_id)
                folder_dict[core_file_name]["extensions"].append(file_extension)
            else:
                folder_dict.update({core_file_name: {"ids":[file_id], "extensions": [file_extension]}})
        else:
            core_file_name = re.sub(r'(?<=[,])(?=[^\s])', r' ', core_file_name)

        if file_extension == 'zip':
            if core_file_name in patient_folders:
                print_color(f'Folder Already Exists', color='y', output_file=main_log_file)
                new_folder_id = existing_patient_folders.get(core_file_name)[0]

                # for each_id in val:
                GdriveAPI.move_file(file_id=file_id, new_folder_id=new_folder_id)
                scripts = [f'Update folders set New_Files_Imported = TRUE, PDF_File_Processed = null where Folder_ID="{new_folder_id}"']
                run_sql_scripts(engine=engine, scripts=scripts)
            else:
                print_color(f'Folder Does Not Exists', color='r', output_file=main_log_file)
                folder_id = GdriveAPI.create_folder(folder_name=core_file_name, parent_folder=record_input_folder_id)

                GdriveAPI.move_file(file_id=file_id, new_folder_id=folder_id)
                scripts = [f'''insert into folders(`Folder_ID`, `Folder_Name`, `New_Files_Imported`)
                                      values("{folder_id}", "{core_file_name}", True)''']
                run_sql_scripts(engine=engine, scripts=scripts)
        else:
            if core_file_name in patient_folders:
                print_color(f'Folder Already Exists', color='y', output_file=main_log_file)
                new_folder_id = existing_patient_folders.get(core_file_name)[0]

                GdriveAPI.move_file(file_id=file_id, new_folder_id=new_folder_id)
                scripts = [f'Update folders set New_Files_Imported = TRUE, PDF_File_Processed = null  where Folder_ID="{new_folder_id}"']
                run_sql_scripts(engine=engine, scripts=scripts)


def map_files_and_folders_to_google_drive(engine, GsheetAPI):
    sheet_name = "Merged Files - Record Input"
    df = pd.read_sql(f'Select * from merge_process', con=engine)
    row_number = GsheetAPI.get_row_count(sheetname=sheet_name)
    print_color(row_number, color='b')
    df['Is_Single_File'] = df['Is_Single_File'].apply(lambda x: True if x == 1 else "")
    df['Has_Zip_Files'] = df['Has_Zip_Files'].apply(lambda x: True if x == 1 else "")
    df['Zip_File_Unpacked'] = df['Zip_File_Unpacked'].apply(lambda x: True if x == 1 else "")
    df['Index_Page_Created'] = df['Index_Page_Created'].apply(lambda x: True if x == 1 else "")
    df['File_Combined'] = df['File_Combined'].apply(lambda x: True if x == 1 else "")
    df['Folder_Moved_To_Processed_Inputs'] = df['Folder_Moved_To_Processed_Inputs'].apply(lambda x: True if x == 1 else "")
    df['Folder_To_Large_To_Combine'] = df['Folder_To_Large_To_Combine'].apply( lambda x: True if x == 1 else "")

    GsheetAPI.write_data_to_sheet(data=df,sheetname=sheet_name, row_number=2, include_headers=False, clear_data=True)




def move_files_to_parent_folder(root_folder, source_folder, folder, zip_exclusion_list, extension_list,
        extension_exclusion_list, prefix_exclusion_list):
    print_color(root_folder, color='r')

    entries = os.listdir(source_folder)
    folders = []
    files = []
    for each_entry in entries:
        print_color(f'entry: {each_entry}', color='g')
        try:
            folder_files = os.listdir(f'{source_folder}\\{each_entry}')
            print_color(f'folder_files {folder_files}', color='y')
            if len(folder_files) >0:
                folders.append(each_entry)
        except Exception as e:
            print_color(e, color='r')
            files.append(each_entry)
            # pass


    # folders = [x for x in entries if os.path.isdir(x)]
    # files = [x for x in entries if not os.path.isdir(x)]

    print_color(f'folders {folders}', color='y')
    print_color(f'files {files}', color='b')

    exclude_folders = [x for x in folders if x in ['__MACOSX']]
    folders_to_move = [x for x in folders if x not in ['__MACOSX']]
    print_color(f'folders_to_move {folders_to_move}', color='b')

    exclude_files = [x for x in files if x.split(".")[-1].lower() in extension_exclusion_list]
    print_color(f'exclude_files {exclude_files}', color='y')
    starts_with = []
    for prefix in prefix_exclusion_list:
        for each_file in files:
            if each_file.startswith(prefix):
                starts_with.append(each_file)
    files_to_move = [x for x in files if x.split(".")[-1].lower() in extension_list]
    print_color(exclude_files, color='g')

    print_color(f'files_to_move: {files_to_move}', color='r')
    print_color(root_folder, color='r')
    print_color(source_folder, color='r')

    if len(exclude_folders)>0:
        for each_folder in exclude_folders:
            print_color(f'Exclude {each_folder}', color='p')
            zip_exclusion_list.append(each_folder)

    if len(exclude_files) > 0 or len(starts_with) >0:
        folder_detail = source_folder.split(f'{root_folder}\\')[-1]
        print_color(f'Exclude {folder_detail}', color='p')
        zip_exclusion_list.append(folder_detail)
    else:
        if root_folder != source_folder:
            print_color(root_folder, color='r')
            print_color(source_folder, color='r')
            print_color(f'files_to_move: {files_to_move}', color='b')
            for each_file in files_to_move:
                source_path = f'{source_folder}\\{each_file}'
                destination_path = f'{root_folder}\\{each_file}'
                print_color(source_path, color='b')
                print_color(destination_path, color='g')
                shutil.move(source_path, destination_path)
            try:
                os.rmdir(source_folder)
            except Exception as e:
                print_color(e, color='r')

        print_color(f'folders_to_move {folders_to_move}', color='g')
        for each_folder in folders_to_move:
            print_color(each_folder, color='y')
            new_source_folder = f'{source_folder}\\{each_folder}'
            print_color(new_source_folder, color='y')
            move_files_to_parent_folder(root_folder, new_source_folder, each_folder, zip_exclusion_list,
                    extension_list, extension_exclusion_list, prefix_exclusion_list)
            # break

    print_color(zip_exclusion_list, color='g')

    return zip_exclusion_list



def unpack_child_folders(GdriveAPI, parent_folder, processed_folder_id, child_folders, extension_exclusion_list, prefix_exclusion_list, patient_log_file):
    excluded_folders = []
    for each_folder in child_folders:
        child_folder_id = each_folder.get("id")
        child_folder_name = each_folder.get("name")
        sub_child_folders = GdriveAPI.get_child_folders(folder_id=child_folder_id)
        if len(sub_child_folders) >0:
            unpack_child_folders(GdriveAPI, parent_folder, processed_folder_id, sub_child_folders)
        child_folder_files = GdriveAPI.get_files(folder_id=child_folder_id)
        child_folder_extensions = [x.get("file_extension").lower() for x in child_folder_files]

        overlapping_extensions = [x for x in extension_exclusion_list if x in child_folder_extensions]
        starts_with = []
        for prefix in prefix_exclusion_list:
            for name in [x.get("name") for x in child_folder_files]:
                if  name.startswith(prefix):
                    starts_with.append(name)
        if len(overlapping_extensions) >0 or len(starts_with) >0:
            print_color(f'Will not unpack Folder because Exclusion file exists', color='r', output_file=patient_log_file)
            excluded_folders.append(each_folder)
        else:
            print_color(child_folder_files, color='p', output_file=patient_log_file)
            for each_child_file in child_folder_files:
                child_file_id = each_child_file.get("id")
                GdriveAPI.move_file( file_id=child_file_id, new_folder_id=parent_folder)

            try:
                GdriveAPI.delete_folder(file_id=child_folder_id, folder_name=child_folder_name)
            except:
                print_color(f'Could Not Delete Folder. WIll Move instead', color='r', output_file=patient_log_file)
                GdriveAPI.move_file( file_id=child_folder_id, new_folder_id=processed_folder_id)
    return excluded_folders

def process_zip_files(GdriveAPI, file_export, folder_id, processed_folder_id, file_id, file_name, extended_file_name, viewable_files, folder_files,
                      extension_list, extension_exclusion_list, prefix_exclusion_list, patient_log_file
                      ):
    numbers = number_list()
    print_color(f'Run Zip Process', color='y', output_file=patient_log_file)
    print_color(file_name, color='r', output_file=patient_log_file)

    core_file_name = ".".join(file_name.split(".")[:-1])
    number_assignment = None
    # file_id = each_file.get("id")
    if core_file_name[-2:].strip() in numbers:
        number_assignment = core_file_name[-2:].strip()
        core_file_name = core_file_name[:-2]
        print_color(core_file_name, color='r', output_file=patient_log_file)
        core_file_name = re.sub(r'(?<=[,])(?=[^\s])', r' ', core_file_name)

    else:
        core_file_name = re.sub(r'(?<=[,])(?=[^\s])', r' ', core_file_name)

    print_color(number_assignment, color='p', output_file=patient_log_file)
    existing_files = [x.get("name") for x in folder_files]
    print_color(existing_files, color='y', output_file=patient_log_file)

    GdriveAPI.download_file(file_id=file_id, file_name=extended_file_name)

    extended_unzipped_folder = f'{file_export}\\{".".join(file_name.split(".")[:-1])} unzipped'
    print_color(extended_unzipped_folder, color='y', output_file=patient_log_file)
    create_folder(extended_unzipped_folder)

    check_zip_files = os.listdir(extended_unzipped_folder)
    print_color(check_zip_files, color='p', output_file=patient_log_file)

    with zipfile.ZipFile(extended_file_name, 'r') as z:
        # Replace 'extracted_folder_id' with the desired folder ID for the extracted contents
        for i, filename in enumerate(z.namelist()):
            sanitized_name = filename.replace(':', '_').replace("/","\\").strip()
            if " \\" in sanitized_name:
                sanitized_name =  sanitized_name.replace(" \\","\\")
            extracted_path = os.path.join(extended_unzipped_folder, sanitized_name)

            normalized_path = os.path.normpath(extracted_path.rstrip())
            print_color(normalized_path, color='b', output_file=patient_log_file)
            paths_to_create = normalized_path.split(extended_unzipped_folder)[-1].split("\\")[1:-1]
            for i in range(len(paths_to_create)):
                create_path = "\\".join(paths_to_create[:i+1])
                create_folder(f'{extended_unzipped_folder}\\{create_path}')
            # z.extract( file_info.filename, extended_unzipped_folder)
            print_color(sanitized_name, color='r', output_file=patient_log_file)
            if sanitized_name[-1] == "\\":
                create_folder(normalized_path)
            else:
                print_color(filename, color='p', output_file=patient_log_file)
                with z.open(filename) as source, open(normalized_path, 'wb') as target:
                    shutil.copyfileobj(source, target)

        # z.extractall(extended_unzipped_folder)
        print_color(f'Files Unzipped into {extended_unzipped_folder}', color='b', output_file=patient_log_file)
    zip_exclusion_list =[]
    zip_exclusions = move_files_to_parent_folder(extended_unzipped_folder,extended_unzipped_folder, extended_unzipped_folder, zip_exclusion_list,
                      extension_list, extension_exclusion_list, prefix_exclusion_list)
    extracted_files = os.listdir(extended_unzipped_folder)
    files_to_upload = [x for x in extracted_files if x.split(".")[-1].lower() in extension_list]
    print_color(extracted_files, color='y', output_file=patient_log_file)
    # for each_item in extracted_files:
    #     if os.path.isdir(f'{extended_unzipped_folder}\\{each_item}'):
    #         shutil.rmtree(f'{extended_unzipped_folder}\\{each_item}')
    # #
    # # extracted_files = os.listdir(extended_unzipped_folder)
    for i, each_unzipped_file in enumerate(files_to_upload):
        file_path = f'{extended_unzipped_folder}\\{each_unzipped_file}'
        extension = each_unzipped_file.split(".")[-1]
        if number_assignment is not None:
            new_file_name = f'{".".join(each_unzipped_file.split(".")[:-1])} {number_assignment}.{extension}'
        else:
            new_file_name = each_unzipped_file
        if new_file_name in existing_files:
            print_color(f'File Already Uploaded', color='r', output_file=patient_log_file)
        else:
            GdriveAPI.upload_file(folder_id=folder_id, file_name=new_file_name, file_path=file_path)

    if len(zip_exclusions)==0:
        delete_success = GdriveAPI.delete_file(file_id=file_id, file_name=file_name)
        if delete_success is False:
            GdriveAPI.move_file(file_id=file_id, new_folder_id=processed_folder_id)

        time.sleep(3)
    # else:

    # os.remove(extended_file_name)
    # # shutil.rmtree(extended_file_name)
    # shutil.rmtree(extended_unzipped_folder)
    return zip_exclusions


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


def sort_image_files(image_files):
    sorted_files = sorted(image_files, key=lambda x: (int(x['size'])), reverse=True)
    return sorted_files

def bytes_to_gb(bytes_value):
    gb_value = bytes_value / (1024 ** 3)
    return gb_value


def get_file_size(sorted_files, extension_list, patient_log_file):
    numbers = number_list()
    print_color(sorted_files, color='b', output_file=patient_log_file)
    data = {
        'ID': [],
        'File Name': [],
        'File Extension': [],
        'File Size': []
    }

    considered_file_df = pd.DataFrame(data)
    print_color(considered_file_df, color='b', output_file=patient_log_file)

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

            print_color(file_name, file_size, file_id, file_ext, color='y', output_file=patient_log_file)

    considered_file_df['File Size'] = considered_file_df['File Size'].astype(int)
    considered_file_df['File Size in GB'] = considered_file_df['File Size'].apply(lambda x: bytes_to_gb(x))

    print_color(considered_file_df, color='b', output_file=patient_log_file)
    print_color(considered_file_df['File Size in GB'].dtype, color='b', output_file=patient_log_file)

    combined_files_size = considered_file_df['File Size in GB'].sum()
    # combined_files_size = bytes_to_gb(files_size)

    print_color(combined_files_size, color='p', output_file=patient_log_file)

    return combined_files_size


def create_index(html_path, pdf_path, folder_name, file_list, excluded_file_list):
    body = f"<h1>{folder_name}</h1><br>"

    body += f'''<span style="color:Black;font-weight:Bold; text-indent: 20px">Combined files:</span><br>'''
    body += f'<table>'

    # index_content = "Index Page\n\n"
    # # body += f'''<br><br><span style="color:Black;font-weight:Bold; ">Player Name:</span> {player_name}'''
    start_page = 0
    page_count = 0
    for i, file_info in enumerate(file_list):
        file_name, start_page, page_count, size, created_time = file_info
        print_color(created_time, color='b')
        date = datetime.datetime.fromisoformat(created_time.replace("Z", "+00:00")).strftime("%m/%d/%Y")
        adjusted_size = round(bytes_to_gb(int(size)) * 1000,2)
        if start_page+2 == start_page +page_count +1:
            body += f'''<tr>
              <td style="font-weight:bold; vertical-align: text-top; width: 30px; text-indent: 10px">{i+1}.</td>
              <td style="font-weight:bold; vertical-align: text-top; width: 150px">Page {start_page + 2}</td>
              <td style="vertical-align: text-top; width: 400px">{file_name}</td>
              <td style="vertical-align: text-top; width: 50px">{adjusted_size} MB</td>
              <td style="vertical-align: text-top; width: 100px">{date}</td>
              </tr>'''

        else:
            body += f'''<tr>
           <td style="font-weight:bold; vertical-align: text-top; width: 50px; text-indent: 30px">{i+1}.</td>
           <td style="font-weight:bold; vertical-align: text-top; width: 150px;"> Pages {start_page+2}-{start_page +page_count +1 }</td>
           <td style="vertical-align: text-top; width: 400px">{file_name}</td>
           <td style="vertical-align: text-top; width: 100px">{adjusted_size} MB</td>
           <td style="vertical-align: text-top; width: 100px">{date}</td>
           </tr>'''

    body += f'''<tr>
             <td style="font-weight:bold; vertical-align: text-top; width: 50px; text-indent: 30px"></td>
        </tr><tr> 
        <td style="font-weight:bold; vertical-align: text-top; width: 50px; text-indent: 30px"> {len(file_list)}</td>
        <td style="font-weight:bold; vertical-align: text-top; width: 150px;">Combined Files</td>
        <td style="font-weight:bold; vertical-align: text-top; width: 400px;">Total Pages {start_page +page_count +1}</td>
    </tr>'''
    body += f'</table>'

    body += f'''<br><br><span style="color:Black;font-weight:Bold; text-indent: 20px">Additional files:</span><br>'''
    body += f'<table>'
    for i, file_info in enumerate(excluded_file_list):
        file_name, size, created_time = file_info
        print_color(created_time, color='b')
        if created_time is not None:
            date = datetime.datetime.fromisoformat(created_time.replace("Z", "+00:00")).strftime("%m/%d/%Y")
        else:
            date = ''
        if size is not None:
            adjusted_size = round(bytes_to_gb(int(size)) * 1000, 2)
        else:
            adjusted_size = ''
        if start_page + 2 == start_page + page_count + 1:
            body += f'''<tr>
                <td style="font-weight:bold; vertical-align: text-top; width: 30px; text-indent: 10px">{i + 1}.</td>
                <td style="vertical-align: text-top; width: 150px"></td>
                <td style="vertical-align: text-top; width: 400px">{file_name}</td>
                <td style="vertical-align: text-top; width: 50px">{adjusted_size} MB</td>
                <td style="vertical-align: text-top; width: 100px">{date}</td>
                </tr>'''

        else:
            body += f'''<tr>
             <td style="font-weight:bold; vertical-align: text-top; width: 30px; text-indent: 10px">{i + 1}.</td>
             <td style="vertical-align: text-top; width: 150px"></td>
             <td style="vertical-align: text-top; width: 400px">{file_name}</td>
             <td style="vertical-align: text-top; width: 100px">{adjusted_size} MB</td>
             <td style="vertical-align: text-top; width: 100px">{date}</td>
             </tr>'''



    if len(excluded_file_list) != 1:
        body += f'''
            <tr>
             <td style="font-weight:bold; vertical-align: text-top; width: 50px; text-indent: 30px"></td>
            </tr>
            <tr> 
                <td style="font-weight:bold; vertical-align: text-top; width: 50px; text-indent: 30px"> {len(excluded_file_list)}</td>
                <td style="font-weight:bold; vertical-align: text-top; width: 150px;">Excluded Files</td>
            </tr>'''

    elif  len(excluded_file_list) == 1:
        body += f'''
            <tr>
             <td style="font-weight:bold; vertical-align: text-top; width: 50px; text-indent: 30px"></td>
            </tr><tr> 
                <td style="font-weight:bold; vertical-align: text-top; width: 50px; text-indent: 30px"> {len(excluded_file_list)}</td>
                <td style="font-weight:bold; vertical-align: text-top; width: 150px;">Excluded File</td>
            </tr>'''
    body += f'</table>'

    with open(html_path, 'w') as f:
        f.write(body)
        f.close()

    converter.convert(html_path, pdf_path)


    return start_page + page_count


def merge_to_pdf(GdriveAPI, sorted_files, excluded_files, folder_exclusions, export_folder_name, folder_name, extension_list,
                       processed_folder_id, record_input_folder_id, folder_id, patient_log_file):
    pdf_writer = PyPDF2.PdfWriter()
    combined_pdf_file = f'{export_folder_name}\\{folder_name} Combined Draft.pdf'
    if os.path.exists(combined_pdf_file):
        os.remove(combined_pdf_file)

    # merger = PdfMerger()
    merger = fitz.open()
    merged_files = []
    file_list = []
    start_page = 0

    for each_file in sorted_files:
        file_id = each_file.get("id")
        file_name = each_file.get("name").replace(":","-")
        core_file_name = each_file.get("core_file_name")
        file_extension = each_file.get("file_extension").lower()
        size = each_file.get("size")
        created_time =  each_file.get("created_time")
        if file_extension in extension_list:
            print_color(size, color='y', output_file=patient_log_file)
            export_file_name = f'{export_folder_name}\\{file_name}'
            print_color(export_file_name, color='y', output_file=patient_log_file)
            if file_extension in ["jpg", "jpeg", "png"]:
                if int(size) < 1000:
                    print_color(f'Image is less than 1 KB so will ignore file', color='r', output_file=patient_log_file)
                    return

            GdriveAPI.download_file(file_id=file_id, file_name=export_file_name)
            print_color(export_file_name, color='g', output_file=patient_log_file)
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
                print_color(f'File Already Exists and is a duplicate. Will not merge', color='r', output_file=patient_log_file)
            else:
                merged_files.append(file_details)
                print_color(f'File Merged into Combined PDF', color='g', output_file=patient_log_file)
                if file_extension != "pdf":
                    converted_export_file_name = f'{".".join(export_file_name.split(".")[:-1])}.pdf'
                    if file_extension in ['doc', 'docx']:
                        convert(export_file_name, converted_export_file_name)
                    elif file_extension in ['jpg', 'jpeg', 'png']:
                        convert_image_to_pdf(export_file_name, converted_export_file_name)
                    adjusted_export_file_name = converted_export_file_name
                else:
                    adjusted_export_file_name = export_file_name

                pdf_document = fitz.open(adjusted_export_file_name)
                merger.insert_pdf(pdf_document)

                # merger.append(adjusted_export_file_name)

                file_list.append((file_name, start_page, page_count, size, created_time))
                start_page += page_count

            GdriveAPI.move_file(file_id=file_id, new_folder_id=processed_folder_id)

    excluded_file_list = []
    for each_file in excluded_files:
        file_id = each_file.get("id")
        file_name = each_file.get("name")
        core_file_name = each_file.get("core_file_name")
        file_extension = each_file.get("file_extension")
        size = each_file.get("size")
        created_time = each_file.get("created_time")
        excluded_file_list.append((file_name, size, created_time))
    for each_folder in folder_exclusions:
        excluded_file_list.append((each_folder, None, None))


    html_path = f'{export_folder_name}\\index.html'
    pdf_path = f'{export_folder_name}\\index.pdf'

    print_color(len(file_list), color='y', output_file=patient_log_file)

    if len(file_list) >0:
        page_count = create_index(html_path, pdf_path, folder_name, file_list, excluded_file_list)
        print_color(merger, color='y', output_file=patient_log_file)
        # merger.write(combined_pdf_file)
        merger.save(combined_pdf_file)
        merger.close()

        final_combined_pdf_file = f'{export_folder_name}\\{folder_name} Combined {page_count}.pdf'
        if os.path.exists(final_combined_pdf_file):
            os.remove(final_combined_pdf_file)

        new_merger = PdfMerger()
        new_merger.append(pdf_path)
        new_merger.append(combined_pdf_file)
        new_merger.write(final_combined_pdf_file)
        new_merger.close()

        print_color('PDF Merged', color='y', output_file=patient_log_file)
        upload_folder_id = record_input_folder_id
        # upload_folder_id = folder_id

        '''check if a combined file already exists. if so, delete'''
        final_upload_file_name = f'{folder_name} Combined {page_count}.pdf'
        all_files = GdriveAPI.get_files(folder_id=record_input_folder_id)
        all_files = [x for x in all_files if x.get("name") == final_upload_file_name]
        if len(all_files) >0:
            GdriveAPI.delete_file(file_id=all_files[0].get("id"), file_name=final_upload_file_name)
        ''' Turned off upload Merge File before OCR'''

        print_color(final_upload_file_name, color='b', output_file=patient_log_file)
        print_color(final_combined_pdf_file, color='y', output_file=patient_log_file)
        print_color(upload_folder_id, color='y', output_file=patient_log_file)

        # GdriveAPI.upload_file(folder_id=upload_folder_id, file_name=final_upload_file_name,
        #                       file_path=final_combined_pdf_file)

        return final_combined_pdf_file, final_upload_file_name

def  ocr_conversion(x, GdriveAPI, upload_folder_id, combined_pdf_file, upload_file_name, patient_log_file):
    ocr_directory = x.ocr_directory
    ocr_settings = x.ocr_setting

    output_directory = combined_pdf_file.split(upload_file_name)[0]
    output_filename = f'{upload_file_name.split(".pdf")[0]} OCR.pdf'
    extended_ouput = f'{output_directory}\\{output_filename}'
    command = f'"{combined_pdf_file}" /output:"{extended_ouput}" /settings:"{ocr_settings}"'
    extended_command = f'"{ocr_directory}\\FileToPDF" {command}'
    print_color(command, color='y', output_file=patient_log_file)
    print_color(extended_command, color='b', output_file=patient_log_file)
    # result = subprocess.run([ocr_directory, ocr_settings], text=True)
    # os.system(extended_command)
    result = subprocess.run(extended_command)
    print_color(result, color='g', output_file=patient_log_file)

    all_files = GdriveAPI.get_files(folder_id=upload_folder_id)
    all_files = [x for x in all_files if x.get("name") == output_filename]
    if len(all_files) > 0:
        GdriveAPI.delete_file(file_id=all_files[0].get("id"), file_name=output_filename)

    GdriveAPI.upload_file(folder_id=upload_folder_id, file_name=output_filename,
                          file_path=extended_ouput)



# @error_handler
def process_individual_folder(x, engine, i, record_input_folder_id, processed_inputs_folder_id, df, merge_process_df,
                              file_export, GsheetAPI, GdriveAPI, extension_list, extension_exclusion_list,
                              prefix_exclusion_list, images_extension_list, patient_log_file):

    folder_id = df['Folder_ID'].iloc[i]
    folder_name = df['Folder_Name'].iloc[i].strip()
    new_files_imported = df['New_Files_Imported'].iloc[i]
    zip_files_exists = df['Zip_Files_Exists'].iloc[i]
    zip_files_unzipped = df['Zip_Files_Unzipped'].iloc[i]
    pdf_file_processed = df['PDF_File_Processed'].iloc[i]

    export_folder_name = f'{file_export}\\{folder_name}'
    create_folder(export_folder_name)

    print_color(f'{i}/{df.shape[0]} Getting files for {folder_id}: {folder_name}', output_file=patient_log_file)
    folder_files = GdriveAPI.get_files(folder_id)
    get_folder_sub_folders = GdriveAPI.get_child_folders(folder_id=folder_id)
    folder_files = [x for x in folder_files if x.get("trashed") == False]
    zip_files = [x for x in folder_files if "zip" in x.get("name").lower()]
    viewable_files = [x for x in folder_files if x.get("name").split(".")[-1].lower() in extension_list]

    has_zip = True if len(zip_files) > 0 else False

    folder_ids = merge_process_df[(merge_process_df['Folder_ID'] == folder_id)]
    if folder_ids.shape[0] > 0:
        scripts = [f'''update merge_process set  
                    Import_Date = curdate(),
                    Is_Single_File= False, 
                    Has_Zip_Files = {has_zip},
                    Zip_File_Unpacked = False, 
                    Index_Page_Created = False, 
                    File_Combined = False, 
                    Folder_Moved_To_Processed_Inputs = False
                    where Folder_ID ="{folder_id}"
                    ''']
    else:
        scripts = [f'''insert into merge_process
                          values(null, curdate(), "{folder_name}", "{folder_id}", "https://drive.google.com/file/d/{folder_id}", False, {has_zip}, False, False, False, False, False)
                                                                           ''']
    run_sql_scripts(engine=engine, scripts=scripts)

    print_color(folder_files, color='y', output_file=patient_log_file)
    print(f'Folder Count {len(folder_files)}', f'Processed Folder Count {len(get_folder_sub_folders)}')
    if len(folder_files) == 0 and len(get_folder_sub_folders) == 0:
        GdriveAPI.delete_folder(folder_id=folder_id, folder_name=folder_name)
        scripts = []
        scripts.append(f'Delete from folders where Folder_ID = "{folder_id}"')
        scripts.append(f'Delete from merge_process where Folder_ID = "{folder_id}"')

        run_sql_scripts(engine=engine, scripts=scripts)

    elif len(folder_files) == 1 and len(get_folder_sub_folders) == 0 and len(zip_files) == 0 and len(
            viewable_files) == 1:
        '''MOVE FILE OUT AS SINGLE FILE / REMOVE FOLDER'''
        ''' RENAME FILE TO CORE LOGIC'''
        each_file = folder_files[0]
        file_id = each_file.get("id")
        file_extension = each_file.get("name").split(".")[-1]

        print_color(folder_name, color='g', output_file=patient_log_file)
        new_file_name = f'{folder_name}.{file_extension}'
        GdriveAPI.rename_file(file_id=file_id, new_file_name=new_file_name)
        GdriveAPI.move_file(file_id=file_id, new_folder_id=record_input_folder_id)
        GdriveAPI.delete_folder(folder_id=folder_id, folder_name=folder_name)
        scripts = []
        scripts.append(f'Delete from folders where Folder_ID = "{folder_id}"')
        scripts.append(f'Delete from merge_process where Folder_ID = "{folder_id}"')
        run_sql_scripts(engine=engine, scripts=scripts)

    else:
        print_color(folder_files, color='y', output_file=patient_log_file)
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
        '''Step 10 - OCR Process'''
        '''Step 11 - Upload Data to Google Sheets'''

        print_color('''Step 1 - Check if there are files that we already processed''', color='r', output_file=patient_log_file)
        processed_folder_id = None
        all_images_folder_id = None
        processed_images_folder_id = None
        inaccessible_files_folder_id = None
        child_folders = GdriveAPI.get_child_folders(folder_id=folder_id)
        folder_exclusions = []

        child_folders = [x for x in child_folders if x.get("trashed") is False]
        print_color(child_folders, color='y', output_file=patient_log_file)
        for each_folder in child_folders:
            if "Processed Files" in each_folder.get("name"):
                processed_folder_id = each_folder.get("id")

            if "All Images" in each_folder.get("name"):
                all_images_folder_id = each_folder.get("id")

            if "Processed Images" in each_folder.get("name"):
                processed_images_folder_id = each_folder.get("id")

            if "Inaccessible Files" in each_folder.get("name"):
                inaccessible_files_folder_id = each_folder.get("id")

        if processed_folder_id is None:
            processed_folder_id = GdriveAPI.create_folder(folder_name='Processed Files', parent_folder=folder_id)

        '''Step 2 - Check if there are folders in the folder. If so unpack files into main folder'''
        print_color('''Step 2 - Check if there are folders in the folder. If so unpack files into main folder''', color='r', output_file=patient_log_file)
        additional_child_folders = [x for x in child_folders if x.get("name") != 'Processed Files' \
                                    and x.get("name") != 'Processed Images'
                                    and x.get("name") != 'All Images']
        print_color(additional_child_folders, color='y', output_file=patient_log_file)
        excluded_folders = unpack_child_folders(GdriveAPI=GdriveAPI, parent_folder=folder_id,
                                                processed_folder_id=processed_folder_id,
                                                child_folders=additional_child_folders,
                                                extension_exclusion_list=extension_exclusion_list,
                                                prefix_exclusion_list=prefix_exclusion_list,
                                                patient_log_file=patient_log_file)
        excluded_folder_name = [x.get("name") for x in excluded_folders]
        folder_exclusions.extend(excluded_folder_name)
        print_color(excluded_folders, color='y', output_file=patient_log_file)


        '''Step 3 - Unzip any Zip Files *'''
        print_color('''Step 3 - Unzip any Zip Files *''',color='r', output_file=patient_log_file)

        for each_file in folder_files:
            file_extension = each_file.get("file_extension")
            file_id = each_file.get("id")
            # pprint.pprint(GdriveAPI.get_file_data(file_id))
            file_name = each_file.get("name")
            extended_file_name = f'{file_export}\\{file_name}'
            print_color(each_file, color='g', output_file=patient_log_file)

            if file_extension == 'zip':
                zip_exclusions = process_zip_files(GdriveAPI, file_export, folder_id, processed_folder_id, file_id,
                                                   file_name, extended_file_name,
                                                   viewable_files, folder_files, extension_list,
                                                   extension_exclusion_list, prefix_exclusion_list, patient_log_file)
                folder_exclusions.extend(zip_exclusions)
                scripts = [f'''update merge_process set Zip_File_Unpacked = True where Folder_ID ="{folder_id}" ''']
                run_sql_scripts(engine=engine, scripts=scripts)

        '''Step 4 - For Files Already Process get file content'''
        print_color('''Step 4 - For Files Already Process get file content''', color='r', output_file=patient_log_file)
        processed_folder_files = GdriveAPI.get_files(processed_folder_id)
        processed_folder_files = [x for x in processed_folder_files if x.get("trashed") == False]

        '''Step 5 - Get Files from Original Folder with New Zipped Files'''
        print_color( '''Step 5 - Get Files from Original Folder with New Zipped Files''', color='r', output_file=patient_log_file)
        folder_files = GdriveAPI.get_files(folder_id)
        folder_files = [x for x in folder_files if x.get("trashed") == False]
        print_color(folder_files, color='b', output_file=patient_log_file)
        if len(folder_files) == 0 and len(processed_folder_files) == 0:
            print_color(f'No New Files to Process', color='r', output_file=patient_log_file)
            return
        # print_color(folder_exclusions, color='r')
        '''Step 6 - Combine Processed and Unprocessed Files'''
        print_color('''Step 6 - Combine Processed and Unprocessed Files''', color='r', output_file=patient_log_file)
        folder_files = folder_files + processed_folder_files
        folder_files = [x for x in folder_files if x.get("trashed") == False]
        excluded_files = [x for x in folder_files if x.get("file_extension").lower() not in extension_list and x.get(
            "file_extension").lower() != "zip"]
        inaccessible_files = [x for x in folder_files if x.get("name").startswith("._")]
        folder_files = [x for x in folder_files if x.get("file_extension").lower() in extension_list]
        folder_files = [x for x in folder_files if not x.get("name").startswith("._")]
        updated_folder_files = [x for x in folder_files if x.get("file_extension").lower() not in images_extension_list]
        image_files = [x for x in folder_files if x.get("file_extension").lower() in images_extension_list]
        print_color(image_files, color='p', output_file=patient_log_file)
        # '''Step 3.5 Move All Images to Image Folder / Move top 50 sorted by file Size to Processed Folder'''
        sorted_image_files = sort_image_files(image_files)

        print_color(f'inaccessible_files {len(inaccessible_files)}', color='y', output_file=patient_log_file)
        if len(inaccessible_files) > 0:
            if inaccessible_files_folder_id is None:
                inaccessible_files_folder_id = GdriveAPI.create_folder(folder_name='Inaccessible Files',
                                                                       parent_folder=folder_id)
            for k, each_file in enumerate(inaccessible_files):
                print_color(f'{k}/{len(inaccessible_files)}', color='g')
                GdriveAPI.move_file(file_id=each_file.get("id"), new_folder_id=inaccessible_files_folder_id)

        print_color(f'Image Files {len(sorted_image_files)}', color='g', output_file=patient_log_file)
        if len(sorted_image_files) > 0:
            if len(sorted_image_files) > 50:
                if all_images_folder_id is None:
                    all_images_folder_id = GdriveAPI.create_folder(folder_name='All Images', parent_folder=folder_id)
                for k, each_image in enumerate(sorted_image_files[50:]):
                    print_color(f'{k}/{len(sorted_image_files[50:])}', color='g', output_file=patient_log_file)
                    GdriveAPI.move_file(file_id=each_image.get("id"), new_folder_id=all_images_folder_id)

            if processed_images_folder_id is None:
                processed_images_folder_id = GdriveAPI.create_folder(folder_name='Processed Images',
                                                                     parent_folder=folder_id)
            for k, each_image in enumerate(sorted_image_files[:50]):
                print_color(f'{k}/{len(sorted_image_files[:50])}', color='g', output_file=patient_log_file)
                GdriveAPI.move_file(file_id=each_image.get("id"), new_folder_id=processed_images_folder_id)

        '''Step 7 - Sort File By File Number - Create Date'''
        print_color('''Step 7 - Sort File By File Number - Create Date''', color='r', output_file=patient_log_file)

        sorted_files = sort_files(updated_folder_files)
        print_color(sorted_files, color='y', output_file=patient_log_file)

        '''Step 8 - Get combined size of all unique files'''
        print_color('''Step 8 - Get combined size of all unique files''', color='r', output_file=patient_log_file)
        combined_files_size = get_file_size(sorted_files, extension_list, patient_log_file)

        '''Step 9 - Merge Files To one PDF'''
        print_color( '''Step 9 - Merge Files To one PDF''', color='r', output_file=patient_log_file)
        # print_color(len(sorted_files),color='y')
        if combined_files_size > .80:
            print_color(f'Combined File Size in folder exceed Allowed Size to run', color='r', output_file=patient_log_file)
            scripts = [f'''update merge_process set Folder_To_Large_To_Combine = True where Folder_ID ="{folder_id}" ''']
            run_sql_scripts(engine=engine, scripts=scripts)
        else:
            final_combined_pdf_file, final_upload_file_name = merge_to_pdf(GdriveAPI, sorted_files, excluded_files,
                       folder_exclusions, export_folder_name, folder_name,extension_list, processed_folder_id,
                                       record_input_folder_id, folder_id, patient_log_file)
            print_color(final_combined_pdf_file, color='g', output_file=patient_log_file)
            print_color(final_upload_file_name, color='y', output_file=patient_log_file)
            ocr_conversion(x, GdriveAPI, record_input_folder_id, final_combined_pdf_file, final_upload_file_name, patient_log_file)


            print_color(folder_id, color='g', output_file=patient_log_file)
            print_color(processed_inputs_folder_id, color='g', output_file=patient_log_file)
            scripts = []
            scripts.append(
                f'Update folders set PDF_File_Processed = True, New_Files_Imported=null where Folder_ID = "{folder_id}"')
            if len(folder_exclusions) > 0 or len(excluded_files) > 0:
                scripts.append(f'''update merge_process set
                                               Index_Page_Created=True,
                                               File_Combined = True,
                                               Folder_Moved_To_Processed_Inputs = FALSE
                                               where Folder_ID ="{folder_id}"''')
            else:
                GdriveAPI.move_file(file_id=folder_id, new_folder_id=processed_inputs_folder_id)

                scripts.append(f'''update merge_process set
                            Index_Page_Created=True,
                            File_Combined = True,
                            Folder_Moved_To_Processed_Inputs = True
                            where Folder_ID ="{folder_id}"''')
            run_sql_scripts(engine=engine, scripts=scripts)

        '''Step 10 - Update Google Sheet'''
        print_color('''Step 10 - Update Google Sheet''', color='r', output_file=patient_log_file)
        map_files_and_folders_to_google_drive(engine, GsheetAPI)



def process_open_folders(x, engine, GdriveAPI, GsheetAPI, record_input_folder_id, processed_inputs_folder_id, main_log_file, merge_log_output_folder):
    print_color( '''PROCESS FOLDERS THAT NEED TO BE MERGED TO A PDF''', color='k', output_file=main_log_file)
    patient_log_output_folder = f'{merge_log_output_folder}\\Patient Logs'
    create_folder(patient_log_output_folder)

    numbers = number_list()
    file_export = f'{x.project_folder}\\Record Inputs'
    create_folder(file_export)
    extension_list = ["pdf", "docx", "doc", "png", "jpeg", "jpg"]
    extension_exclusion_list = ["exe", "bat", "dll"]
    images_extension_list = [ "png", "jpeg", "jpg"]
    prefix_exclusion_list = ["._"]

    df = pd.read_sql(f'''Select * from folders where (New_Files_Imported is null or New_Files_Imported = 1)
        and (PDF_File_Processed != 1 or PDF_File_Processed is null)
        and Folder_Name not in ("1 - Folders For Review With Alan", "Processed Inputs", "Doubt Files", "Old Reports", "Repeat Files")
--         and Folder_Name in ("2024.03.20, Robilotti, Cecilia")
        order by Folder_Name
    ''', con=engine)
    merge_process_df = pd.read_sql(f'Select * from merge_process', con=engine)

    print_color(df, color='r', output_file=main_log_file)

    for i in range(df.shape[0]):
        now = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        folder_name = df['Folder_Name'].iloc[i]
        patient_log_file = f'{patient_log_output_folder}\\{folder_name} {now}.html'

        try:
            process_individual_folder(x, engine, i, record_input_folder_id, processed_inputs_folder_id, df, merge_process_df,
                                  file_export, GsheetAPI, GdriveAPI, extension_list, extension_exclusion_list,
                                  prefix_exclusion_list, images_extension_list, patient_log_file)
        except Exception as e:
            print_color(e, color='r', output_file=patient_log_file)
        break

def merge_files_to_pdf(x, environment):
    log_output_folder = x.log_output_folder
    create_folder(log_output_folder)
    merge_log_output_folder = f'{log_output_folder}\\Merge Logs'
    create_folder(merge_log_output_folder)

    now = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    main_log_file = f'{merge_log_output_folder}\\Merge Logs {now}.html'
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
    print_color(f'folder_id: {folder_id}', color='b', output_file=main_log_file)
    child_folders = GdriveAPI.get_child_folders(folder_id=folder_id)
    for each_folder in child_folders:
        if each_folder.get("name") == 'RECORD-INPUT':
            record_input_folder_id = each_folder.get("id")
            break

    print_color(f'Record Input folder ID {record_input_folder_id}', color='y', output_file=main_log_file)
    sub_child_folders = GdriveAPI.get_child_folders(folder_id=record_input_folder_id)
    sub_child_folders = [x for x in sub_child_folders if x.get("trashed") is False]
    print_color(f'Record Input Folders {sub_child_folders}', color='r', output_file=main_log_file)

    processed_folder_id = None
    for each_folder in sub_child_folders:
        if each_folder.get("name") == 'Processed Inputs':
            processed_folder_id = each_folder.get("id")
            sub_child_folders.remove(each_folder)
            break

    print_color(f'Processed Folder ID {processed_folder_id}', color='y', output_file=main_log_file)
    run_sql_scripts(engine=engine_1, scripts=[f'Truncate folders;'])

    '''GET DICT OF ALL FOLDERS IN RECORD-INPUT'''

    existing_patient_folders = get_existing_patient_folders(GdriveAPI, record_input_folder_id, sub_child_folders, processed_folder_id, main_log_file=main_log_file)
    '''RENAME FOLDERS THAT ARE NOT FORMATTED PROPERLY'''

    rename_existing_folders(GdriveAPI, existing_patient_folders, main_log_file)
    '''GET UPDATED DICT OF ALL FOLDERS IN RECORD-INPUT'''
    existing_patient_folders = get_existing_patient_folders(GdriveAPI, record_input_folder_id, sub_child_folders, processed_folder_id, include_processed_folders=True, main_log_file=main_log_file)
    '''MERGE FOLDERS THAT HAVE THE SAME NAME'''
    merge_existing_folders(GdriveAPI, existing_patient_folders, record_input_folder_id, main_log_file)
    '''GET UPDATED DICT OF ALL FOLDERS IN RECORD-INPUT'''
    existing_patient_folders = get_existing_patient_folders(GdriveAPI, record_input_folder_id, sub_child_folders, processed_folder_id, main_log_file=main_log_file)
    '''MAP FOLDERS TO SQL'''
    import_new_folders(engine_1, database_name, existing_patient_folders, main_log_file)
    '''GET UPDATED DICT OF ALL FOLDERS IN RECORD-INPUT'''
    existing_patient_folders = get_existing_patient_folders(GdriveAPI, record_input_folder_id, sub_child_folders, processed_folder_id, include_processed_folders=True, main_log_file=main_log_file)
    '''MAP / MOVE NEW FILES IN RECORD INPUT'''
    process_new_files(engine_1, GdriveAPI, record_input_folder_id, existing_patient_folders, main_log_file)
    '''PROCESS FOLDERS THAT NEED TO BE MERGED TO A PDF'''
    process_open_folders(x, engine_1, GdriveAPI, GsheetAPI, record_input_folder_id, processed_folder_id, main_log_file, merge_log_output_folder)
    '''MAP MERGE PROCESS TO GOOGLE SHEETS'''





