#### Choose a folder directory and extract text from all Invoice pdfs inside,
#### Find invoice infos and paste them into a CSV file

import os
import fitz # This is pymupdf
import logging
import csv
import itertools
import datetime
from re import compile
from tkinter import Tk
from tkinter.filedialog import askdirectory

# For logging and debugging
logging.basicConfig(level=logging.DEBUG,format=' %(asctime)s - %(levelname)s - %(message)s')
logging.debug('Start of program')

# A class for invoice elements
class inv_element:
    inv_element_list = []
    def __init__(self, name, pattern):
        self.name   = name
        self.pattern  = pattern
        self.__class__.inv_element_list.append(self)

    # Compile regex object and return the first non-null group in matching result
    def getResult(self,text):
        regex = compile(self.pattern)
        group_count = self.pattern.count('|') + 1
        try:
            match = regex.search(text)
            return  match.group(1)
        except AttributeError:
            return None 

# Function to get all text from a pdf
def getAllText(path):
    doc = fitz.open(path)
    text = ''
    for page in doc:
        text += page.get_text()
    return text

### Create invoice elements 
inv_number = inv_element('Invoice Number',r'Invoice No.*\n(.*)|Invoice Number\:(\d*)')
inv_date = inv_element('Invoice Date',r'Invoice Date\:\n?(.*)')
inv_curr = inv_element('Invoice Currency',r'Amount\((\w*)')
inv_desc = inv_element('Invoice Description',r'Amount\(.*\)\n.\n(.*)|Amount\(.*\)\n(.*)')
amt_beforeTax = inv_element('Amount before Tax',r'Amount\(.*\)\n.*\n(\d+.*)|(\d+.*)\nOutput Tax')
amt_tax = inv_element('Amount of Tax',r'Output Tax\s?\n(.*)')
amt_afterTax  = inv_element('Amount after Tax',r'Total Amount in\s?\(?.*\)?\n?(.*)')
po_number = inv_element('PO Number', r'PO Number.*(.*)|PO No.*\n?(.*)')
buyer_name = inv_element('Buyer Name',r'BILL TO PARTY.*\n(.*)|To\n(.*)')
buyer_taxcode = inv_element(' Buyer Tax code', r'VAT No\s?\:(.*)')
seller_name = inv_element('Seller Name', r'REGISTERED\s?\n?OFFICE\s?.?\n?(.*)')
seller_taxcode = inv_element('Seller Tax code',r'VAT NO\s?\:(.*)')

### Get input prompt to select folder directory
Tk().withdraw() 
folderPath = askdirectory() 

## Get all the pdf files in selected folder
name_list = [pdf for pdf in os.listdir(folderPath) if pdf.upper().endswith('.PDF')]
inv_list = {}

### Create nested dict with each sub dict is the pdf
for i in name_list:
    inv_list[i] = {}
    inv_list[i]['Path'] = os.path.join(folderPath,i) #Get path
    inv_list[i]['Text'] = getAllText(inv_list[i].get('Path'))
    
    # A loop to iterate through list of text and list of regex
    for element in inv_element.inv_element_list:  
        inv_list[i][element.name] = element.getResult(inv_list[i].get('Text'))


### Write invoice data list into a csv
## Get header for csv
headers_csv = ['File Name','Path','Text']
for element in inv_element.inv_element_list:  
    headers_csv.append(element.name)

## Get CSV name at time created
dt = datetime.datetime.now()
csv_dt = dt.strftime('%Y%m%d %H.%M.%S')

## Write CSV and save it at user selected folder
with open(f'{folderPath}/output {csv_dt}.csv','w',newline='',encoding="utf-16") as output_csv:
    w = csv.DictWriter(output_csv, headers_csv, delimiter = '\t')
    w.writeheader()
    for key,val in sorted(inv_list.items()):
        row = {'File Name':key}
        row.update(val)
        w.writerow(row)
output_csv.close()

logging.debug('End of program')
