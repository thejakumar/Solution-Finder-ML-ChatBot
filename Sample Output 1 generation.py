import time
import mammoth
import os
import codecs
from bs4 import BeautifulSoup
import re
import xlsxwriter

z = os.getcwd()
f = open(z+'\\SampleInputDoc2-.docx','rb')
b = open('x.html','wb')
document = mammoth.convert_to_html(f)
b.write(document.value.encode('utf8'))
f.close()
b.close()

r=codecs.open("x.html", 'r').read()

soup = BeautifulSoup(r,"lxml")
company_name = soup.find_all('strong')
company_name1 = soup.find_all('h3')

lst = list()
asd = company_name[0]
a = re.sub("<.*?>", "", asd.text)
lst.append(a)
for i in range (0,len(company_name1)):
    x = company_name1[i]
    a = re.sub("<.*?>", "", x.text)
    lst.append(a)

ad = []
for i in range (1,16):
            a = company_name[i]
            a = re.sub("<.*?>", "", a.text)
            ad.append(a)
az = "If you're having problems loading up Windows Explorer and browsing your file system, the problem is almost always a shell extension that shouldn't be installed, or some shell extensions that are conflicting with each other. For example, the shell extensions for Dropbox and TortoiseSVN tend to cause problems when you put your code into your Dropbox folder, causing hanging and generally slow file browsing.Your best bet is to grab a copy of ShellExView and start disabling third-party shell extensions, or uninstalling Windows Explorer plug-ins that you don't actually need. You can also use this tool in combination with ShellMenuView to clean up your messy Explorer context menu."
ad.append(az)

workbook = xlsxwriter.Workbook('SampleOutput1.xlsx')
worksheet = workbook.add_worksheet()

row = 0
col = 0

for i in range (0,4):
    worksheet.write(row, col, lst[i])
    row += 1

row = 0
col = 1

worksheet.write(row, col, ad[0]+". "+ad[1]+". "+ad[2]+"."+ad[3]+".")
row+=1

worksheet.write(row, col, ad[4]+". "+ad[5]+". "+ad[6]+"."+ad[7]+"."+ad[8]+". "+ad[9]+"."+ad[10]+".")
row+=1

worksheet.write(row, col, ad[11]+". "+ad[12]+". "+ad[13]+"."+ad[14]+".")
row+=1

worksheet.write(row, col, az)
workbook.close()

print("Sample Output 1 Generated,.....")
time.sleep(2)
