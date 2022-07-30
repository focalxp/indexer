from __future__ import print_function
import PySimpleGUI as sg
import os.path
       
from PyPDF2 import PdfFileReader, PdfFileWriter
from pathlib import Path
import re
import PyPDF2

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2 import service_account


SERVICE_ACCOUNT_FILE = 'keys.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
creds = None
creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1DQyjh1-8TEFKUqqSi9b-TNcjsCtnCspWiMbv3t2mJeM'


service = build('sheets', 'v4', credentials=creds)

# Call the Sheets API
sheet = service.spreadsheets()
result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                            range="Sheet1!B1:B508").execute()
result2 = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                            range="Sheet1!A1:A508").execute()
values = result.get('values')
values2 = result2.get('values')

name = []
nameraw = []
for x in values:
        for i in x:
                name.append(i)

for x in values2:
        for i in x:
                nameraw.append(i)

sg.theme("DarkTeal2")
layout = [[sg.T("")], [sg.Text("Choose a file: "), sg.Input(), sg.FileBrowse(key="-IN-")],[sg.Button("RUN")]]

###Building Window
window = sg.Window('My File Browser', layout, size=(600,150))

def indexer (file_path):
    pdf = PdfFileReader(file_path)
    pagesfinal = [[]]
    for i in range(len(name)):
            print(nameraw[i])
            pagesfinal.append([])
            for page_num in range (pdf.numPages):    
                pageobj = pdf.getPage(page_num)
                pageinfo = pageobj.extractText()
                pageinfo = ''.join(pageinfo.split())
                name[i]= ''.join(str(name[i]).split())
                nameraw[i] = ''.join(str(nameraw[i]).split())
                if (re.search(name[i], pageinfo)) or (re.search(nameraw[i], pageinfo)) or (re.search(name[i].upper(), pageinfo)):
                        pagesfinal[i].append(page_num +1)
                        print(pagesfinal[i])
        
            request = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="Sheet1!C1", valueInputOption = "USER_ENTERED", body={"values":pagesfinal}).execute()
        

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event=="Exit":
        break
    elif event == "RUN":
        indexer(values["-IN-"])

