from textwrap import indent
from tkinter import *
import tkinter
from tkinter.ttk import *
from tkinter.filedialog import askopenfile, askopenfilenames
import os
import pandas
import openpyxl
root = Tk()
root.geometry('1000x500') #window resolution

def getExcelData(file_name):
    if os.path.splitext(file_name)[1]=='.xlsx':
        return pandas.read_excel(file_name).to_string(index=False)
    
    if os.path.splitext(file_name)[1]=='.csv' :
        df= pandas.read_csv(file_name)
        return df.to_string(index=False)
    return ''
       
def open_file():

    def updateUI(data, cnt, dict_d):
        # T.insert(root,data)
        T.insert(tkinter.END,data)
        T1.insert(tkinter.END, cnt)
        T2.insert(tkinter.END, dict_d)

    files = askopenfilenames(filetypes =[('Excel Files', '*.xlsx'),('CSV','*.csv')])
    if files is not None:
        s=''
        count=0
        dict_d=''
        for file_name in files:
            count+=1
            s+=os.path.basename(file_name)+'\n'
            if (getExcelData(file_name)!= ''):
                dict_d+=os.path.basename(file_name)+'\n\n'+getExcelData(file_name)+'\n\n\n\n\n\n'
            
        updateUI(str(s), count, dict_d)


    



btn = Button(root, text ='Browse file', command = lambda:open_file())
btn.pack(side = TOP, pady = 10)

T = Text(root, height = 3, width = 50)
l = Label(root, text = "Files selected")
T1 = Text(root, height = 3, width = 50)
T2 = Text(root, height = 400, width = 1000, font=('',9))
l1= Label(root, text = "Count")

l.pack()
T.pack()
l1.pack()
T1.pack()
T2.pack()
mainloop()

