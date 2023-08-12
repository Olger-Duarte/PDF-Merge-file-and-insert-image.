# !/usr/bin/python

import re
import os
import ocrmypdf
import fitz
import tkinter as tk
import pypdfium2 as pdfium
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.lib import utils
from reportlab.pdfgen import canvas
from tkinter import filedialog
from PyPDF2 import PdfWriter, PdfReader, PageObject

global error_list
error_list = []

def extract_data_from_label(text):
    client_name = []
    sku_code = []
    data_dict = {}

    print(f"Found essential data...")
    for i in range(len(text)):
        #Client name
        if 'SHIP TO' in text[i]:
            remaining_name = text[i+1].strip().title()
            client_name.append(remaining_name)
        elif text[i].startswith('SHIP '):
            remaining_name = text[i].split("SHIP", 1)[-1].strip().title()
            client_name.append(remaining_name)
        
        #SKU Code
        remaining_code = re.findall(r'\d{4} \d{4} \d{4} \d{4} \d{4} \d{2}', text[i])
        if remaining_code:
            sku_code.append(remaining_code[0])
    
    data_dict = dict(zip(client_name, sku_code))
    return data_dict

#Generic read PDF function
def read_pdf(label_file_path):
    print(f"Reading file {os.path.basename(label_file_path)}...")
    try:
        with open(label_file_path, "rb") as pdf_file:
            pdf_reader = PdfReader(pdf_file)
            final_text = ""
            for page in pdf_reader.pages:
                final_text += page.extract_text()
            return final_text
    except Exception as e:
        final_error = f"Error reading file: {str(e)}"
        error_list.append(final_error) 
        print(final_error)
    
#OCR
def ocr_pdf(label_file_path):
    fle_name = os.path.basename(label_file_path)
    print(f"Using OCR {fle_name}...\n")
    try:
        output_ocr_pdf = f"OCR_{fle_name}"  # Tempoeral file
        ocrmypdf.ocr(
            label_file_path, output_ocr_pdf,
            output_type='pdf', skip_text=True, deskew=True
        )
        return read_pdf(output_ocr_pdf)
    except Exception as e:
        final_error = f"Error proccesing '{os.path.basename(label_file_path)}': {str(e)}"
        error_list.append(final_error) 
        print(final_error)
        return None
    finally:
        # Eliminate teporal OCR File
        if output_ocr_pdf:
            try:
                os.remove(output_ocr_pdf)
            except Exception as e:
                print(f"Error eliminating temporal file '{output_ocr_pdf}': {str(e)}")

#Obtain Label as images for insert in new PDF file
def convert_pdf_to_image(pdf_path, image_path):
    print(f"Converting label into images...")
    image_name = []
    try:
        pdf = pdfium.PdfDocument(pdf_path)
        version = pdf.get_version()  # get the PDF standard version
        nums_pages = len(pdf)  # get the number of pages in the document

        page_indices = [i for i in range(nums_pages)]  # all pages
        renderer = pdf.render(
            pdfium.PdfBitmap.to_pil,
            page_indices = page_indices,
            scale = 600/72,  # 600dpi resolution
        )
        for i, image in zip(page_indices, renderer):
            image_name.append(f"{image_path}\out_label_{nums_pages}_{i}.jpg")
            image.save(image_name[i])
    except Exception as e:
        final_error = f"Error convert PDF to image: {str(e)}"
        error_list.append(final_error) 
        print(final_error)
    return image_name


def main():
    #PDF directories:
    default_folder_current = os.path.dirname(os.path.abspath(__file__))
    default_folder_labels = f"{default_folder_current}\PDF_Input\Label"
    default_folder_booklets = f"{default_folder_current}\PDF_Input\Booklets"
    default_folder_additionals = f"{default_folder_current}\PDF_Input\Additionals"
    default_folder_imag_labels = f"{default_folder_current}\PDF_Input\Stickers"
    default_folder_sku = f"{default_folder_current}\PDF_Output"
    print(f"Actual project folder {default_folder_current}")
    root = tk.Tk()
    root.withdraw()

    #Selected file
    print(f"Select shipping label file")
    label_file_path = filedialog.askopenfilename(
        title="Select shipping label file", 
        filetypes=[("PDF Files", "*.pdf")],
        initialdir=default_folder_labels
    )
    print(f"Select brochure PDF file")
    booklet_repo_path = filedialog.askopenfilename(
        title="Seleccione el archivo PDF de folletos",
        filetypes=[("PDF Files", "*.pdf")],
        initialdir=default_folder_booklets
    )
    print(f"Select PDF Brochure extra pages file")
    additional_pages_path = filedialog.askopenfilename(
        title="Seleccione el archivo de páginas adicionales de folletos PDF", 
        filetypes=[("PDF Files", "*.pdf")],
        initialdir=default_folder_additionals
    )

    #If the file cannot be read, "Optical character recognition" is tried in order to identify the text.
    final_text = read_pdf(label_file_path)
    if final_text is None or "" or len(final_text) == 0:
        final_text = ocr_pdf(label_file_path)
    
    #The subtracted text is entered into a list to analyze each element and subtract the essential data.
    all_text = final_text.split("\n")
    data_of_label = extract_data_from_label(all_text)
    #Get Client name
    keys_list = list(data_of_label.keys())
    client_name = keys_list[0].split()
    first_name = client_name[0]
    last_name = client_name[1]
    #Get SKU code
    values_list = list(data_of_label.values())
    sku_code = values_list[0]
    #Get actual date
    current_date = datetime.now().strftime("%Y-%m-%d")
    
    # Converting the label PDF to an image
    images_names = convert_pdf_to_image(label_file_path, default_folder_imag_labels)

    #Create SKU and insert booklet pages and merge additional pages
    try:
        #Creat TEMP file
        temp_first_pages = f"{default_folder_sku}\TEMP_FISRT.pdf"  # Tempoeral file
        temp_last_pages = f"{default_folder_sku}\TEMP_LAST.pdf" 

        # Extract first two pages of the booklet
        with open(temp_first_pages, "wb") as first_pages:
            pdf_writer = PdfWriter()
            with open(booklet_repo_path, "rb") as booklet_pdf_file:
                print(f"Extracting fisrt booklet pages...")
                booklet_pdf_reader = PdfReader(booklet_pdf_file)
                for page_num in range(2):
                    pdf_writer.add_page(booklet_pdf_reader.pages[page_num])
            pdf_writer.write(first_pages)

        # Extract last two pages of the booklet
        with open(temp_last_pages, "wb") as last_pages:
            pdf_writer = PdfWriter()
            with open(booklet_repo_path, "rb") as booklet_pdf_file:
                print(f"Extracting last booklet pages...")
                booklet_pdf_reader = PdfReader(booklet_pdf_file)
                for page_num in range(-2,0):
                    pdf_writer.add_page(booklet_pdf_reader.pages[page_num])
            pdf_writer.write(last_pages)
        
        #Merge additional pages
        print(f"Mergeing additional pages...")
        merger = PdfWriter()
        for pdf in [temp_first_pages, additional_pages_path,]:
            merger.append(pdf)
        temp_marge = f"{default_folder_sku}\TEMP_MERGE.pdf"
        merger.write(temp_marge)

        for pdf in [temp_marge, temp_last_pages,]:
            merger.append(pdf)
        output_sku_file = f"{default_folder_sku}\SKU.pdf"
        merger.write(output_sku_file)
        print(f"Finish SKU file: {output_sku_file}")

        #Insert image label
        print("Inserting image label...")
        image_path = images_names[0]
        #output_sku_file = f"{default_folder_sku}/SKU.pdf"
        image = open(image_path, "rb").read()
        doc = fitz.open(output_sku_file)
        page = doc[0]
        page_width = page.rect.width
        page_height = page.rect.height
        x1 = (page_width/100) * 62
        y1 = (page_height/100) * 48
        x2 = (page_width/100) * 98
        y2 = (page_height/100) * 98
        image_rectangle = fitz.Rect(x1, y1, x2, y2) # (0, Up/Down, right/left, 0)    
        page.insert_image(image_rectangle, stream=image)
        
        #Create final name of PDF
        final_file_name = f"{last_name}-{sku_code.replace(' ', '')}-{current_date}"
        doc.save(f"{default_folder_sku}\{final_file_name}.pdf")

        #Eliminate temp files
        os.remove(temp_first_pages)
        os.remove(temp_last_pages)
        os.remove(temp_marge)

        print(f"Done!: {final_file_name}")
    except Exception as e:
        final_error = f"Error formatting the document: {str(e)}"
        error_list.append(final_error) 
        print(final_error)
  
    log_file_path = os.path.join(os.path.dirname(label_file_path), f"LOG_{last_name}-{sku_code.replace(' ', '')}.txt")
    with open(log_file_path, "a") as error_file:
        for error in error_list:
            error_file.write(error + '\n')

if __name__ == "__main__":
    main()
