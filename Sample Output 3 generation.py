import xlrd 
import pandas as pd
import textwrap
import array
import time
import xlrd
import os
import docx2txt
from textblob import TextBlob
from nltk import BlanklineTokenizer
import re

z = os.getcwd()
workbook = xlrd.open_workbook(z+'\\SampleInput.xlsx')
sheets = workbook.sheet_names()
required_data = []
for sheet_name in sheets:
    sh = workbook.sheet_by_name(sheet_name)
    for rownum in range(sh.nrows):
        row_values = sh.row_values(rownum)
        required_data.append((row_values[4]))
required_data2 = []
for sheet_name in sheets:
    sh = workbook.sheet_by_name(sheet_name)
    for rownum in range(sh.nrows):
        row_values = sh.row_values(rownum)
        required_data2.append((row_values[5]))        
required_data1 = list(filter(None, required_data))

z = os.getcwd()
text = docx2txt.process(z+"\\SampleInputDoc1-FAQs.docx")
blob = TextBlob(text)
tokenizer = BlanklineTokenizer()
z = blob.tokenize(tokenizer)
c = '?'
lst = list()
for i in range (0,len(z)):
    x = z[i].find(c)
    if x!= -1:
        lst.append(z[i])

import xlsxwriter

workbook = xlsxwriter.Workbook('SampleOutput3.xlsx')
worksheet = workbook.add_worksheet()

row = 0
col = 0

for problem in (required_data):
    worksheet.write(row, col, problem)
    row += 1

row = 0
col = 1

for solution in (required_data2):
    worksheet.write(row, col, solution)
    row += 1


row = 400
col = 0

for i in range (0,len(z)):
    if i%2 == 0:
        worksheet.write(row, col, z[i])
        row += 1

row = 400
col = 1

for i in range (0,len(z)):
    if i%2 != 0:
        worksheet.write(row, col, z[i])
        row += 1

workbook.close()

print("Sample Output Generated,.....")
time.sleep(2)

orders = pd.read_excel("SampleOutput3.xlsx",0)
orders.drop([428,430,431,432,433,434,435,436,437,438,439,440,441], axis  = 0, inplace = True)
writer = pd.ExcelWriter("SampleOutput3.xlsx")
orders.to_excel(writer, index = False)
writer.save()
