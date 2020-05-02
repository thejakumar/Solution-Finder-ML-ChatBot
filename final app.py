import xlrd
import os
from tkinter import *
from tkinter import ttk

z = os.getcwd()
workbook = xlrd.open_workbook(z+'\\SampleOutput.xlsx')
sheets = workbook.sheet_names()
required_data = []

for sheet_name in sheets:
    sh = workbook.sheet_by_name(sheet_name)
    for rownum in range(sh.nrows):
        row_values = sh.row_values(rownum)
        required_data.append((row_values[0]))

required_data2 = []

for sheet_name in sheets:
    sh = workbook.sheet_by_name(sheet_name)
    for rownum in range(sh.nrows):
        row_values = sh.row_values(rownum)
        required_data2.append((row_values[1]))
required_data21 = list(filter(None, required_data))

required_data1 = list(set(required_data21[1:]))

top = Tk()
v = Label(top, text="Problem:")
v.grid(row = 0 , column = 0)
def on_keyrelease(event):

    value = event.widget.get()
    value = value.strip().lower()

    def remove_duplicates(values):
        output = []
        seen = set()
        for value in values:
            if value not in seen:
                output.append(value)
                seen.add(value)
        return output

    result = remove_duplicates(required_data1)

    if value == '':
        data = required_data1
    else:
        data = []
        for item in result:
            if value in item.lower():
                data.append(item)
    

    listbox_update(data)


def listbox_update(data):
    listbox.delete(0, 'end')

    data = sorted(data, key=str.lower)

    for item in data:
        listbox.insert('end', item)


def on_select(event):
    global ert
    ert = event.widget.get(event.widget.curselection())
    entry.delete(0,END)
    entry.insert(0,ert)

entry = ttk.Entry(top, width=100)
entry.bind('<KeyRelease>', on_keyrelease)
ff = ttk.Scrollbar(top)
fff = ttk.Scrollbar(top,orient = "horizontal")
listbox = Listbox(top, width = 100 , yscrollcommand=ff.set, xscrollcommand = fff.set)
ff.grid(column=1,row=2,sticky="NS")
fff.grid(column=0,row=3,sticky="EW")
ff.config(command = listbox.yview)
fff.config(command = listbox.xview)
entry.grid(row = 1, column = 0)
listbox.grid(column=0,row=2)
listbox.bind('<<ListboxSelect>>', on_select)
listbox_update(required_data1)
def z():
    y.delete('1.0',END)
def close_window ():
    z = entry.get()
    for sheet in workbook.sheets():
        for rowidx in range(sheet.nrows):
            row = sheet.row(rowidx)
            for colidx, cell in enumerate(row):
                if cell.value == z :
                    aov = sheet.cell_value(rowidx,colidx+1)
                    aov1 = aov.split("<br />")
                    a = len(aov1)
                    lst=[]
                    for i in range(0,a):
                        lst.append(len(aov1[i]))
                    max_val = lst[0]       
                    for item in lst:       
                        if item > max_val: 
                            max_val = item
                    for i in range(0,a):
                        if lst[i] == max_val:
                            global y
                            y.insert(END,aov1[i]+"\n")
def zz():
    top = Tk()
    v = Label(top, text="Your Gmail address:")
    v.pack()
    st = Entry(top)
    st.pack()
    v1 = Label(top, text="Your Gmail Password:")
    v1.pack()
    st1 = Entry(top,show="*")
    st1.pack()
    v2 = Label(top, text="Query:")
    v2.pack()
    st2 = Entry(top)
    st2.pack()
    def z():
        import smtplib
        s = smtplib.SMTP('smtp.gmail.com', 587)
        s.starttls()
        s.login(st.get(), st1.get())
        message = st2.get()
        s.sendmail(st.get(), "thotapoornithsundar@gmail.com", message)
        s.quit()
        top.destroy()
        y.insert(END,"Query Sent.....")
    v = ttk.Button(top, text = "LOGIN", command = z)
    v.pack()
    top.geometry("500x300")
    top.title("Email Feedback")
    top.mainloop()
w = ttk.Button(top, text="CHECK",command = close_window)
w.grid(row = 4)
ww = ttk.Button(top, text="CLEAR",command = z)
ww.grid(row = 5)
www = ttk.Button(top, text="Problem not listed?",command = zz)
www.grid(row = 7)
f = ttk.Scrollbar(top)
y = Text(top, wrap = WORD, yscrollcommand=f.set)
y.grid(column = 0,row = 6)
f.grid(column = 1,row = 6, sticky='NS')
f.config(command = y.yview)
top.title("ABC Corp - Solution Tracer")
top.mainloop()
