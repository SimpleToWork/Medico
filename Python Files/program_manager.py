import os
from zipfile import ZipFile
from pypdf import PdfMerger
from global_modules import print_color, create_folder

def unzip_files(zip_path, unzip_path):
    with ZipFile(zip_path, 'r') as zObject:
        zObject.extractall(path=unzip_path)

    print_color(f'Zip Folder Extracted', color='g')



def merge_files_to_pdf():
    zip_path = f'C:\\Users\\Ricky\\Downloads\\2023.08.07, Beth Strumpf 3.zip'
    unzip_path =  f'C:\\Users\\Ricky\\Downloads\\2023.08.07, Beth Strumpf 3'
    pdf_file = f'{unzip_path}\\pdf_sample.pdf'
    create_folder(unzip_path)
    unzip_files(zip_path, unzip_path)

    files_in_folder = os.listdir(unzip_path)
    files_in_folder = [x for x in files_in_folder if ".pdf" in x.lower()]
    print_color(files_in_folder, color='y')

    # pdfs = ['file1.pdf', 'file2.pdf', 'file3.pdf', 'file4.pdf']

    merger = PdfMerger()

    for pdf in files_in_folder:
        merger.append(f'{unzip_path}\\{pdf}')
        print_color(f'{pdf} Merged', color='b')

    merger.write(pdf_file)
    merger.close()

    print_color('PDF Merged', color='y')

def run_program():
    merge_files_to_pdf()



if __name__ == '__main__':
    run_program()
