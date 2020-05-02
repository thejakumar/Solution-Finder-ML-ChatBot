import mammoth
import os
import re
import time
import codecs
from bs4 import BeautifulSoup
import re
import xlsxwriter

z = os.getcwd()
f = open(z+'\\SampleInputDoc3-Hardware Problems.docx','rb')
b = open('xy.html','wb')
document = mammoth.convert_to_html(f)
b.write(document.value.encode('utf8'))
f.close()
b.close()

r=codecs.open("xy.html", 'r').read()

soup = BeautifulSoup(r,"lxml")
company_name2 = soup.find_all('ul')
company_name1 = soup.find_all('h4')
company_name = soup.find_all('p')
headings = []
h1 = []
lines = []
for i in range (1,len(company_name2)):
            global aaaa
            aaaa = company_name2[i]
            aaaaa = re.sub("<.*?>", "", aaaa.text)
            lines.append(aaaaa)
for i in range (1,len(company_name)):
            global a
            a = company_name[i]
            aa = re.sub("<.*?>", "", a.text)
            headings.append(aa)
for i in range (1,len(company_name1)):
            global aaa
            aaa = company_name1[i]
            aaaaaa = re.sub("<.*?>", "", aaa.text)
            h1.append(aaaaaa)

ans1 = lines[0]
ques1 = "Unresponsive PC"
ques2 = headings[1]
for i in range (1,10):
    ans2 = lines[i]
ques3 = headings[11]
ans3 = lines[10]
ques4 = headings[12]
ans4 = lines[11]
ques5 = headings[13]
ans5 = lines[12]
ques6 = headings[14]
ans6 = lines[13]
ques7 = headings[15]
ans7 = lines[14]
ques8 = headings[17]
ans8 = lines[15]
ques9 = headings[18]
ans9 = lines[16]
ques10 = headings[19]
ans10 = lines[17]
ques11 = headings[20]
ans11 = lines[18]
ques12 = headings[21]
ans12 = lines[19]
ans12_1 = headings[22]
ans13 = lines[20]
ques14 = headings[23]
ans14_1 = headings[25]
ans14 = lines[21]
ques15 = headings[26]
ans15 = lines[22]
ques16 = headings[27]
ans16 = lines[23]
ques17 = headings[29]
ans17_1 = headings[30]
ans17_2 = lines[24]
ans17_3 = headings[31]
ans17_4 = lines[25]
ques18 = headings[34]
ans18 = headings[35]
ques19 = headings[36]
ans19 = lines[26]
ques20 = headings[37]
ans20 = headings[38]
ques21 = headings[43]
ans21 = headings[44]
ans21_1 = lines[27]
ans21_2 = lines[28]
ques22 = headings[46]
ans22 = lines[29]
ques23 = headings[49]
ans23 = lines[30]
ans23_1 = lines[31]
ques24 = headings[51]
ans24 = lines[32]
ques25 = headings[55]
ans25 = headings[56]
ans25_1 = lines[33]
ans25_2 = headings[57]
ans25_3 = headings[58]
ans25_4 = lines[34]
ans25_5 = headings[59]
ans25_6 = lines[35]
ques26 = headings[60]
ans26 = headings[61]
ans26_1 = lines[36]
ques27 = headings[62]
ans27 = lines[37]
ques28 = headings[63]
ans28 = lines[38]
ques29 = headings[64]
ans29 = lines[39]
ques30 = headings[65]
ans30 = lines[40]
ques31 = headings[66]
ans31 = lines[41]
ques = []
ans = []

ques.extend((ques1, ques2, ques3, ques4, ques5, ques6, ques12, ques14, ques17, ques18, ques19, ques20, ques21, ques22, ques23, ques24, ques25, ques26, ques27, ques28, ques29, ques30, ques31))
ans.extend((ans1, ans2, ans3, ans4, ans5, ans6+" "+ques7+" "+ans7+" "+ques8+" "+ans8+" "+ques9+" "+ans9+" "+ques10+" "+ans10+" "+ques11+" "+ans11, ans12+" "+ans12_1+" "+ans13, ans14_1+" "+ans14+" "+ques15+" "+ans15+" "+ques16+" "+ans16, ans17_1+" "+ans17_2+" "+ans17_3+" "+ans17_4, ans18, ans19, ans20, ans21+" "+ans21_1+" "+ans21_2, ans22, ans23+" "+ans23_1, ans24, ans25+" "+ans25_1+" "+ans25_2+" "+ans25_3+" "+ans25_4+" "+ans25_5+" "+ans25_6, ans26+" "+ans26_1, ans27, ans28, ans29, ans30, ans31))

workbook = xlsxwriter.Workbook('SampleOutput2.xlsx')
worksheet = workbook.add_worksheet()

row = 0
col = 0

for i in range (0,len(ques)):
    worksheet.write(row, col, ques[i])
    row += 1

row = 0
col = 1

for i in range (0,len(ans)):
    worksheet.write(row, col, ans[i])
    row += 1

workbook.close()


print("Sample Output 2 Generated,.....")
time.sleep(2)
