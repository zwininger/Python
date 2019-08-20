# Import required libraries
import pandas as pd
import numpy as np
import requests
from requests.adapters import HTTPAdapter
from requests.packages.urllib3.util.retry import Retry
from PyPDF2 import PdfFileReader
import re

# Set folder path for 13F documents
folder_path = "S:\\CIO\\Operations\\Middle Office\\Data\\13F\\"

# File location of 13f PDF
path = folder_path + "13f.pdf"

# Create file object for PDF
pdfFileObj = open(path)

# Get number of pages of report
number_of_pages = ''
with open(path, 'rb') as f:
    pdf = PdfFileReader(f)
    number_of_pages = pdf.getNumPages()

# Add all pages to dataframe
thirteen_f = pd.DataFrame()
with open(path, 'rb') as f:
    pdf = PdfFileReader(f)
    for i in range(2,number_of_pages): # First page is title page
        page = pdf.getPage(i)
        text = page.extractText()
        df = pd.Series(text.split('\n'))
        thirteen_f = pd.concat([thirteen_f, df])

# List of rows to remove from dataframe
drop_rows_list = [
    'CUSIP NOISSUER NAMEISSUER DESCRIPTIONSTATUS',
    'IVM001',
    'Run Date',
    'List of Section 13F',
    'Page',
    'Run Time',
    '2019Qtr:',
    'Total Count:'
]

# Remove rows containing values from drop_rows_list
for row in drop_rows_list:
    thirteen_f = thirteen_f[~thirteen_f[0].str.contains(row)]
# Replace blank spaces with NaN
thirteen_f[0].replace('', np.nan, inplace=True)
# Drop NaN values
thirteen_f = thirteen_f.dropna()
# Get first 9 characters from CUSIP column
thirteen_f['CUSIP'] = thirteen_f[0].str[:9]
# Create column for CUSIPs which were deleted from the report
thirteen_f['DELETED'] = thirteen_f[0].str.endswith('DELETED')
# Rename columns
thirteen_f.columns = ['RAW_TEXT', 'CUSIP', 'DELETED']
# Reset index
thirteen_f = thirteen_f.reset_index(drop=True)
# Save 13F file to location
thirteen_f.to_excel(folder_path + "13f.xlsx")