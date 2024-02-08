'''
download tesseract - https://github.com/tesseract-ocr/tesseract/tree/e082522c248d3121e466959a8ba4fd4f7ad1a525
download poppler - https://pdf2image.readthedocs.io/en/latest/installation.html
'''

from pdf2image import convert_from_path
import pytesseract
import PyPDF2

from global_modules import print_color
pytesseract.pytesseract.tesseract_cmd = f'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'
poppler_path = f'C:\\Program Files\\Poppler\\poppler-23.11.0\\Library\\bin'


def read_pdf_page(pdf_path):
    # with open(pdf_path, 'rb') as file:
    pdf_reader = PyPDF2.PdfReader(pdf_path)

    return pdf_reader

def convert_pdf_file_to_ocr(pdf_file):
    pdf_reader = read_pdf_page(pdf_file)
    print(len(pdf_reader.pages))

    # len(pdf_reader.pages)
    counter = 0
    color = 'g'
    for i in range(170, 171):
        if counter < 10:
            if counter % 4 ==0:
                color='r'
            elif counter % 4 ==1:
                color='y'
            elif counter % 4 == 2:
                color = 'g'
            elif counter % 4 == 3:
                color = 'b'

            page = pdf_reader.pages[i].extract_text()
            print_color(page, color=color)
        counter +=1
    # for page_number in range(1,3):


    images = convert_from_path(pdf_file,poppler_path=poppler_path, first_page=171, last_page=171)
    # threshold =2
    # threshold_hash = imagehash.ImageHash(threshold)
    for i, each_image in enumerate(images):
        text = pytesseract.image_to_string(each_image, lang='eng')
        print_color(f"Page {i + 1} Text:\n{text}", color='y')
    #     print_color(each_image, color='y')
    #     img_hash = imagehash.average_hash(each_image)
    #
    #     each_image.show()
    #     print_color(img_hash, color='r')
    #     if img_hash.__lt__(threshold_hash):
    #         print(f"Skipping Page {i + 1} as it appears to be text-based.")
    #         continue
    #



pdf_file = f'C:\\Users\\Ricky\\Desktop\\New Projects\\Medico\\Medico\\Record Inputs\\2024.01.10,  Fitzpatrick, Jean\\2024.01.10,  Fitzpatrick, Jean Combined.pdf'
convert_pdf_file_to_ocr(pdf_file)