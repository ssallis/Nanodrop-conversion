# -*- coding: utf-8 -*-
"""
Created on Sun Dec 29 12:40:20 2019

@author: shawn
"""

import pandas as pd
import numpy as np
import os
from bs4 import BeautifulSoup

from tkinter import Tk
from tkinter.filedialog import askopenfilename

def read_excel_xml(path):
    file = open(path).read()
    soup = BeautifulSoup(file,'xml')
    workbook = []
    for sheet in soup.findAll('Worksheet'): 
        sheet_as_list = []
        for row in sheet.findAll('Row'):
            row_as_list = []
            for cell in row.findAll('Cell'):
                row_as_list.append(cell.Data.text)
            sheet_as_list.append(row_as_list)
        workbook.append(sheet_as_list)
    return workbook

Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file
 
xml_data = read_excel_xml(filename)
df = pd.DataFrame(np.asarray(xml_data[0])[1:,1:], columns=np.asarray(xml_data[0])[0,1:])
dfnew = pd.DataFrame()
dfnew['Sample ID'] = df['Sample ID']
dfnew['User ID'] = df['User name']
dfnew['Date'] = pd.to_datetime(df["Date and Time"]).dt.strftime('%#m/%d/%Y')
dfnew['Time'] = pd.to_datetime(df['Date and Time']).dt.strftime('%#I:%M:%S %p')
dfnew['Conc.'] = pd.to_numeric(df['Nucleic Acid']).round(1)
dfnew['Units'] = pd.DataFrame(['ng/ul']*df.shape[0])
dfnew['A260'] = pd.to_numeric(df['A260 (Abs)']).round(3)
dfnew['A280'] = pd.to_numeric(df['A280 (Abs)']).round(3)
dfnew['260/280'] = pd.to_numeric(df['260/280']).round(2)
dfnew['260/230'] = pd.to_numeric(df['260/230']).round(2)
dfnew['Conc. Factor (ng/ul)'] = pd.to_numeric(df['Factor'])
dfnew['Cursor Pos.'] = pd.DataFrame([260]*df.shape[0])
dfnew['Cursor Abs'] = pd.DataFrame(['']*df.shape[0])
dfnew['340 raw'] = pd.DataFrame(['']*df.shape[0])
dfnew['NA Type'] = pd.DataFrame(['RNA-40']*df.shape[0])

output_filename = os.path.splitext(os.path.basename(filename))[0]
output_path = os.path.split(filename)[0] + '/' + output_filename + '.txt'

dfnew.to_csv(output_path, sep="\t", index=False)
