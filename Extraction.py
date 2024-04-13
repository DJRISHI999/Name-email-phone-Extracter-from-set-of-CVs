
'''You need to write a code which extracts emailID, contact no. and overall text from CV’s and then can be downloaded in .XLS format. 

-	share a working link of your program(not code), where input will be a bundle of CV and output would be excel containing above said information.
-	Sharing code through Git/Email would not be entertained
-	You can use any framework you are comfortable with
Upload your code on any free server and share a working link. Please check if your link is accepting proper input and giving expected output before sharing the link.'''


# Importing required libraries
import os
import re
import pandas as pd
from docx import Document
from tkinter import Tk
import pdfplumber as pr
from tkinter.filedialog import askopenfilename

# ASK USER TO SELECT ZIP FILE

#Asking if user is giving data in zip file or folder
print('Please select the zip file containing the CVs')


Tk().withdraw()
foldername = askopenfilename()
print(foldername)

# Extracting the zip file

import zipfile
if foldername.endswith('.zip'):
    print('Extracting the zip file')
    with zipfile.ZipFile(foldername, 'r') as zip_ref:
        zip_ref.extractall('CVs')

# Function to extract text from docx files
def extract_text_from_docx(file_path):
    doc = Document(file_path)
    text = ''
    for para in doc.paragraphs:
        text += para.text
    return text

# Function to extract text from pdf files
def extract_text_from_pdf(file_path):
    with pr.open(file_path) as pdf:
        text = ''
        for page in pdf.pages:
            text += page.extract_text()
    return text


def extract_text_from_web_edited(file_path):
    try:
        with open(file_path, 'r') as file:
            text = file.read()
    except Exception as e:
        print(f'Error reading {file_path}: {e}')
        text = 'ask candidate to upload file in .docx or .pdf format'
    return text


# Function to extract email,name and phone number from text

def extract_info(text):
    email = re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", text)
    phone = re.findall(r"\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}", text)
  
    return email, phone


#Extracting another folder from extracted folder
if len(os.listdir('CVs')) == 0:
    print('No files found in the folder')
    exit()
# if there is another folder inside the extracted folder, extract the files from that folder
if os.path.isdir(os.path.join('CVs', os.listdir('CVs')[0])):
    folder = os.listdir('CVs')[0]
    folder_path = os.path.join('CVs', folder)
else: 
    folder_path = 'CVs'
    
# Extracting all files from the folder
files = os.listdir(folder_path)

# Creating a dataframe to store the extracted information
data = []
for file in files:
    file_path = os.path.join(folder_path, file)
    print(file_path)
    try:
        if file.endswith('.docx'):
            text = extract_text_from_docx(file_path)

        elif file.endswith('.pdf'):
            text = extract_text_from_pdf(file_path)

        elif file.endswith('.doc'):
            text = extract_text_from_web_edited(file_path)

        else:
            text = ''

    except Exception as e:
        print(f'Error reading {file}: {e}')
        text = ''


        
    email, phone = extract_info(text)
    data.append([file.split('.')[0], email[0] if email else 'cannot decode the file, Ask candidate to send another resume in docx or pdf format', phone[0] if phone else ''])

df = pd.DataFrame(data, columns=['File', 'Email', 'Phone'])
df.to_excel('output.xlsx', index=False)
print('Output saved to output.xlsx')



