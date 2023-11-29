
import re
import os
from zipfile import ZipFile
from pypdf import PdfMerger
import PyPDF2
from docx2pdf import convert

import pandas as pd
from sqlalchemy import inspect
from global_modules import print_color, create_folder, run_sql_scripts, Get_SQL_Types, engine_setup
from google_drive_class import GoogleDriveAPI
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


def get_docx_page_count(file_path):
    doc = Document(file_path)
    page_count = sum(1 for _ in doc.element.xpath('//w:sectPr'))
    return page_count

def get_existing_patient_folders(GdriveAPI, response_folder_id):
    sub_child_folders = GdriveAPI.get_child_folders(folder_id=response_folder_id)
    # print_color(sub_child_folders, color='g')

    # for each_folder in sub_child_folders:
    #     if each_folder.get("name") == 'Processed Inputs':
    #         processed_folder_id = each_folder.get("id")
    #         break

    existing_patient_folders = {}
    for each_folder in sub_child_folders:
        folder_name = each_folder.get('name')
        folder_id = each_folder.get('id')
        # if "2023.11.02, Delusso, Fabienne" in folder_name:
        if folder_name in existing_patient_folders.keys():
            existing_patient_folders[folder_name].append(folder_id)
        else:
            existing_patient_folders.update({folder_name: [each_folder.get('id')]})

    # print_color(existing_patient_folders, color='y')

    return existing_patient_folders


def rename_existing_folders(GdriveAPI, existing_patient_folders):
    numbers = number_list()

    print_color(len(existing_patient_folders.keys()), color='g')
    counter = 0
    for key, val in existing_patient_folders.items():
        core_folder_name = key.strip().replace("  ", " ")
        if core_folder_name[-2:].strip() in numbers:
            core_folder_name = core_folder_name[:-2].strip()

        core_folder_name = re.sub(r'(?<=[,])(?=[^\s])', r' ', core_folder_name)
        if key.strip() != core_folder_name:
            counter +=1
            print_color(core_folder_name, color='p')
            for each_item in val:
                print_color(key, "            ", core_folder_name, each_item, color='y')
                GdriveAPI.rename_folder(folder_id=each_item, new_folder_name=core_folder_name)
    print_color(f'{counter} Folders Renamed', color='g')


def merge_existing_folders(GdriveAPI, existing_patient_folders):
    existing_patient_folders_with_duplicates = {}
    for key, val in existing_patient_folders.items():
        if len(val) > 1:
            existing_patient_folders_with_duplicates.update({key: val})

    print_color(len(existing_patient_folders_with_duplicates.keys()), color='b')
    # print_color(existing_patient_folders_with_duplicates.keys(), color='g')
    counter = 0
    for key, val in existing_patient_folders_with_duplicates.items():
        print_color(f'Folder has Duplicate Entries. Will Merge', color='p')
        original_folder_id = val[0]
        for each_item in val[1:]:
            ''' Get Files in Folder / Move Files / Remove Folder '''
            print_color(each_item, original_folder_id, color='b')
            folder_files = GdriveAPI.get_files(each_item)
            for each_file in folder_files:
                each_file_id = each_file.get("id")

                print_color(each_file, color='y')
                GdriveAPI.move_file(file_id=each_file_id, new_folder_id=original_folder_id)
                counter +=1
            GdriveAPI.delete_folder(folder_id=each_item, folder_name=key)
        # break
    print_color(f'{counter} Folders Merged', color='g')


def import_new_folders(engine, database_name, existing_patient_folders):
    table_name = 'folders'
    if inspect(engine).has_table(table_name):
        data_df = pd.read_sql(f'Select "SQL" as `TYPE`, Folder_ID, Folder_Name from {table_name}', con=engine)
    else:
        data_df = pd.DataFrame()

    print_color(data_df, color='r')
    df = pd.DataFrame.from_dict(existing_patient_folders, orient='columns').transpose()
    df['Folder_Name'] = df.index
    df.insert(0, "TYPE", "Google Drive")
    df = df.reset_index(drop=True)

    df.columns = ["TYPE", 'Folder_ID', 'Folder_Name']
    print_color(df, color='g')
    # df = df[[ 'Folder_ID', 'Folder_Name']]

    df = pd.concat([df, data_df]).drop_duplicates(subset=['Folder_ID', 'Folder_Name'], keep=False)
    df = df[df['TYPE'] == "Google Drive"]
    print_color(df, color='p')

    df = df.drop(columns=['TYPE'])

    sql_types = Get_SQL_Types(df).data_types
    df.to_sql(name=table_name, con=engine, if_exists='append', index=False, schema=database_name, chunksize=1000,
              dtype=sql_types)


def process_new_files(engine, GdriveAPI, response_folder_id, existing_patient_folders):
    numbers = number_list()
    all_files = GdriveAPI.get_files(folder_id=response_folder_id)
    print_color(len(all_files), color='y')
    all_files= [x for x in all_files if "Combined.pdf" not in x.get('name')]
    print_color(len(all_files), color='y')

    folder_dict = {}

    for i, each_file in enumerate(all_files):
        print_color(i, each_file, color='g')
        core_file_name = ".".join(each_file.get("name").split(".")[:-1])
        file_id = each_file.get("id")
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
            print_color(core_file_name, color='y')
            folder_dict.update({core_file_name: [file_id]})

    folder_dict = dict(sorted(folder_dict.items()))
    print_color(folder_dict, color='g')

    new_folder_dict = {}
    for key, val in folder_dict.items():
        if len(val) > 1:
            new_folder_dict.update({key:val})
    print_color(new_folder_dict.keys(), color='y')
    ''' CORE NAME = DATE, LAST NAME, FIRST NAME'''
    ''' ONLY MOVE FILES THAT HAVE MORE THAN 1 FILE PER CORE NAME'''
    patient_folders = list(existing_patient_folders.keys())
    for key, val in new_folder_dict.items():
        print_color(key, val, color='g')
        if key in patient_folders:
            print_color(f'Folder Already Exists', color='y')
            new_folder_id = existing_patient_folders.get(key)[0]

            for each_id in val:
                GdriveAPI.move_file(file_id=each_id, new_folder_id=new_folder_id)
            scripts = [f'Update folders set New_Files_Imported = TRUE where Folder_ID="{new_folder_id}"']
            run_sql_scripts(engine=engine, scripts=scripts)
        else:
            print_color(f'Folder Does Not Exists', color='r')
            folder_id = GdriveAPI.create_folder(folder_name=key, parent_folder=response_folder_id)
            for each_id in val:
                GdriveAPI.move_file(file_id=each_id, new_folder_id=folder_id)
            scripts = [f'''insert into folders(`Folder_ID`, `Folder_Name`, `New_Files_Imported`)
                            values("{folder_id}", "{key}", True)''']
            run_sql_scripts(engine=engine, scripts=scripts)


def process_zip_files(GdriveAPI, file_export, folder_id, file_id, file_name, extended_file_name, viewable_files):
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
    GdriveAPI.delete_file(file_id=file_id, file_name=file_name)

    time.sleep(3)
    os.remove(extended_file_name)
    # shutil.rmtree(extended_file_name)
    shutil.rmtree(extended_unzipped_folder)
    #
    # os.remove(extended_unzipped_folder)

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


def merge_to_pdf(GdriveAPI, sorted_files, export_folder_name, folder_name, extension_list,
                       processed_folder_id, response_folder_id, folder_id):

    combined_pdf_file = f'{export_folder_name}\\{folder_name} Combined.pdf'
    if os.path.exists(combined_pdf_file):
        os.remove(combined_pdf_file)
    merger = PdfMerger()
    merged_files = []
    for each_file in sorted_files:
        file_id = each_file.get("id")
        file_name = each_file.get("name")
        core_file_name = each_file.get("core_file_name")
        file_extension = each_file.get("file_extension")
        size = each_file.get("size")
        if file_extension in extension_list:
            export_file_name = f'{export_folder_name}\\{file_name}'
            GdriveAPI.download_file(file_id=file_id, file_name=export_file_name)

            if file_extension.lower() in ['doc', 'docx']:
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
                    convert(export_file_name, converted_export_file_name)
                    adjusted_export_file_name = converted_export_file_name
                else:
                    adjusted_export_file_name = export_file_name
                merger.append(adjusted_export_file_name)

            GdriveAPI.move_file(file_id=file_id, new_folder_id=processed_folder_id)

    merger.write(combined_pdf_file)
    merger.close()

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
                          file_path=combined_pdf_file)

def process_open_folders(x, engine, GdriveAPI, response_folder_id):

    file_export = f'{x.project_folder}\\Record Inputs'
    create_folder(file_export)
    extension_list = ["pdf", "docx", "doc", "png", "jpeg", "jpg"]
    df = pd.read_sql(f'''Select * from folders where (New_Files_Imported is null or New_Files_Imported = 1)
        and (PDF_File_Processed != 1 or PDF_File_Processed is null)
        and Folder_Name not in ("1 - Folders For Review With Alan", "Processed Inputs", "Doubt Files", "Old Reports")
       -- and Folder_Name in ("2023.11.29, Rodriguez, Doreen")
        order by Folder_Name
    ''', con=engine)
    print_color(df, color='r')

    for i in range(df.shape[0]):
        folder_id = df['Folder_ID'].iloc[i]
        folder_name = df['Folder_Name'].iloc[i]
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
        viewable_files = [x for x in folder_files if x.get("name").split(".")[-1].lower() in extension_list]
        # print_color(folder_files, color='y')
        print(len(folder_files) , len(get_folder_sub_folders))
        if len(folder_files) == 0 and len(get_folder_sub_folders) == 0:
            GdriveAPI.delete_folder(folder_id=folder_id, folder_name=folder_name)
            scripts = []
            scripts.append(f'Delete from folders where Folder_ID = "{folder_id}"')
            run_sql_scripts(engine=engine, scripts=scripts)

        elif len(folder_files) == 1 and len(get_folder_sub_folders) == 0:
            '''MOVE FILE OUT AS SINGLE FILE / REMOVE FOLDER'''
            file_id = folder_files[0].get("id")
            GdriveAPI.move_file(file_id=file_id, new_folder_id=response_folder_id)
            GdriveAPI.delete_folder(folder_id=folder_id, folder_name=folder_name)
            scripts = []
            scripts.append(f'Delete from folders where Folder_ID = "{folder_id}"')
            run_sql_scripts(engine=engine, scripts=scripts)

        else:
            print_color(folder_files, color='y')
            '''Step 1 - Unzip any Zip Files *'''
            '''Step 2 - Check if there are files that we already processed'''
            '''Step 3 - For Files Already Process get file content'''
            '''Step 4 - Get Files from Original Folder with New Zipped Files'''
            '''Step 5 - Combine Processed and Unprocessed Files'''
            '''Step 6 - Sort File By Create Date'''
            '''Step 7 - Merge Files To one PDF
                      - Create Index Page'''
            '''       - Move Files to "Processed Folder"'''

            '''Step 1 - Unzip any Zip Files *'''

            for each_file in folder_files:
                file_extension = each_file.get("file_extension")
                file_id = each_file.get("id")
                # pprint.pprint(GdriveAPI.get_file_data(file_id))
                file_name = each_file.get("name")
                extended_file_name = f'{file_export}\\{file_name}'
                print_color(each_file, color='g')
                if file_extension == 'zip':
                    process_zip_files(GdriveAPI, file_export, folder_id, file_id, file_name, extended_file_name, viewable_files)

            '''Step 2 - Check if there are files that we already processed'''
            processed_folder_id = None
            child_folders = GdriveAPI.get_child_folders(folder_id=folder_id)
            for each_folder in child_folders:
                if "Processed Files" in each_folder.get("name"):
                    processed_folder_id = each_folder.get("id")
                    break

            if processed_folder_id is None:
                processed_folder_id = GdriveAPI.create_folder(folder_name='Processed Files', parent_folder=folder_id)

            '''Step 3 - For Files Already Process get file content'''
            processed_folder_files = GdriveAPI.get_files(processed_folder_id)
            processed_folder_files = [x for x in processed_folder_files if x.get("trashed") == False]

            '''Step 4 - Get Files from Original Folder with New Zipped Files'''
            folder_files = GdriveAPI.get_files(folder_id)
            folder_files = [x for x in folder_files if x.get("trashed") == False]
            print_color(folder_files, color='b')
            if len(folder_files) ==0:
                print_color(f'No New Files to Process', color='r')
                break
            '''Step 5 - Combine Processed and Unprocessed Files'''
            folder_files = folder_files + processed_folder_files
            folder_files = [x for x in folder_files if x.get("trashed") == False]

            '''Step 6 - Sort File By File Number - Create Date'''
            sorted_files = sort_files(folder_files)
            '''Step 7 - Merge Files To one PDF'''
            print_color(len(sorted_files),color='y')
            merge_to_pdf(GdriveAPI, sorted_files, export_folder_name, folder_name, extension_list, processed_folder_id,
                         response_folder_id, folder_id)

            scripts = []
            scripts.append(f'Update folders set PDF_File_Processed = True, New_Files_Imported=null where Folder_ID = "{folder_id}"')
            run_sql_scripts(engine=engine, scripts=scripts)

        # break




        # break


def merge_files_to_pdf(x, environment):
    engine = engine_setup(hostname=x.hostname, username=x.username, password=x.password, port=x.port)
    engine_1 = engine_setup(project_name=x.project_name, hostname=x.hostname, username=x.username, password=x.password,
                            port=x.port)
    database_name =x.project_name
    database_setup(engine=engine, database_name=database_name)
    table_setup(engine_1)


    GdriveAPI = GoogleDriveAPI(credentials_file=x.drive_credentials_file, token_file=x.drive_token_file,
                               scopes=x.drive_scopes)
    folder_id = x.mle_folder
    print_color(folder_id, color='b')
    child_folders = GdriveAPI.get_child_folders(folder_id=folder_id)
    for each_folder in child_folders:
        if each_folder.get("name") == 'RECORD-INPUT':
            response_folder_id = each_folder.get("id")
            break

    print_color(response_folder_id, color='y')

    '''GET DICT OF ALL FOLDERS IN RECORD-INPUT'''
    existing_patient_folders = get_existing_patient_folders(GdriveAPI, response_folder_id)
    '''RENAME FOLDERS THAT ARE NOT FORMATTED PROPERLY'''
    rename_existing_folders(GdriveAPI, existing_patient_folders)
    '''GET UPDATED DICT OF ALL FOLDERS IN RECORD-INPUT'''
    existing_patient_folders = get_existing_patient_folders(GdriveAPI, response_folder_id)
    '''MERGE FOLDERS THAT HAVE THE SAME NAME'''
    merge_existing_folders(GdriveAPI, existing_patient_folders)
    '''GET UPDATED DICT OF ALL FOLDERS IN RECORD-INPUT'''
    existing_patient_folders = get_existing_patient_folders(GdriveAPI, response_folder_id)
    '''MAP FOLDERS TO SQL'''
    import_new_folders(engine_1, database_name, existing_patient_folders)
    '''MAP / MOVE NEW FILES IN RECORD INPUT'''
    process_new_files(engine_1, GdriveAPI, response_folder_id, existing_patient_folders)
    '''PROCESS FOLDERS THAT NEED TO BE MERGED TO A PDF'''
    process_open_folders(x, engine_1, GdriveAPI, response_folder_id)










    # child_folders = GdriveAPI.get_child_folders(folder_id=folder_id)


    # zip_path = f'C:\\Users\\Ricky\\Downloads\\2023.08.07, Beth Strumpf 3.zip'
    # unzip_path =  f'C:\\Users\\Ricky\\Downloads\\2023.08.07, Beth Strumpf 3'
    # pdf_file = f'{unzip_path}\\pdf_sample.pdf'
    # create_folder(unzip_path)
    # unzip_files(zip_path, unzip_path)
    #
    # files_in_folder = os.listdir(unzip_path)
    # files_in_folder = [x for x in files_in_folder if ".pdf" in x.lower()]
    # print_color(files_in_folder, color='y')
    #
    # # pdfs = ['file1.pdf', 'file2.pdf', 'file3.pdf', 'file4.pdf']
    #


