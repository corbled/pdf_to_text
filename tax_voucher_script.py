from PyPDF2 import PdfWriter, PdfReader
import win32com.client
import os
from PIL import Image
from pdf2image import convert_from_path
import pytesseract
import re


# Request pdf filename from user
print("Please insert the pdf name without .pdf: ")
# Folder where individual pdfs will be stored
folder_name = input()
# Original pdf file name variable
full_german_vouch = folder_name + '.pdf'

# If none already exists create folder to store individual pdfs
newpath = 'C:\\Users\\coren\\OneDrive\\Documents\\Python Scripts\\Phil Project\\' + folder_name
if not os.path.exists(newpath):
    os.makedirs(newpath)

# Divide original pdf into individual pages
inputpdf = PdfReader(open(full_german_vouch, "rb"))
for i in range(len(inputpdf.pages)):
    output = PdfWriter()
    output.add_page(inputpdf.pages[i])

    # Name individual pages based off number
    filePath = folder_name + '/document-page%s.pdf' % i
    with open(filePath, "wb") as outputStream:
        output.write(outputStream)

    
    # Insert doc read here
    # For each individual page search for 
    
    # Convert pdf to image
    # file = open(filePath)
    img = convert_from_path(filePath, poppler_path='C:\\Program Files\\poppler\\Library\\bin')
    img[0].save('pic.jpg')

    # for pdf in img:
    #     for i in range(len(pdf)):
    #             img[i].save(f'PDF\image_mods\image_converted_{i+1}.png', 'PNG')
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    txt = pytesseract.image_to_string(Image.open('pic.jpg'), lang='eng', config='--psm 6')
    # # for page_number, page_data in enumerate(img):
    # # txt = pytesseract.image_to_string(img)
    # print(txt)
    instrument_name = re.search(r'(?<=DE\d{10}).*', txt)
    print(f'Instrument Name = {instrument_name}')
    # if instrument_name != "None":
    #     print(f'Instrument Name = {instrument_name.group(0)}')
    

        # print("Page # {} - {}".format(str(page_number),txt))


    # stock_name = 

    # Rename individual pdf based off position name
    # old_name = 'individual_pdfs/document-page%s.pdf' % i
    # new_name = stockName
    # os.rename(old_name, new_name)





### Email sending portion, commented out while testing OCR ####

# ## Send email ##
# ol = win32com.client.Dispatch('Outlook.Application')
# olmailitem = 0x0
# newmail = ol.CreateItem(olmailitem)

# # Email content
# newmail.Subject = 'Testing Mail'
# newmail.To = 'cbled@interactivebrokers.ie'
# newmail.Body = 'This is a test'

# # Add attachment
# for i in range(len(inputpdf.pages)):
#     attach = 'C:\\Users\\coren\\OneDrive\\Documents\\Python Scripts\\Phil Project\\' + folder_name + '\\document-page%s.pdf' % i
#     newmail.Attachments.Add(attach)

# # newmail.Display()
# newmail.Send()


# TODO
    # Move the original pdf into the relevant folder
    # os.rename('C:\\Users\\cbled\\OneDrive - Interactive Brokers Group, Inc\\Documents\\Phil German tax vouchers\\' + full_german_vouch, 'C:\\Users\\cbled\\OneDrive - Interactive Brokers Group, Inc\\Documents\\Phil German tax vouchers\\' + folderName + '\\' + full_german_vouch)
    # Add pdf image reader functionality. For each individual pdf created
        # Read image and store all text in image
        # Query image for the postion name and date
        # Rename file to position -  clientName_positionName_Year - eg, Tamis_Volkswagen_2022.pdf
    # Test sending emails to cadocs
    # Create exe that would run on Phils machine, would need to have the precise directory location


