import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from tkinter import *
from tkinter import ttk

#convert report to a csv file
def toCSV(file):
    pd.read_excel(file, 'Page 1', index_col=None).to_csv('report.csv', encoding='utf-8')

#create dataframe
def createDF(filename, config):
    df = pd.read_csv(filename, delimiter=',', usecols=['Number', 'Configuration item', 'Category', 'Subcategory', 'Created',
                                                       'Resolved', 'Opened by', 'Assignment group'])
    #check for Datamart/ETL tickets & remove if found
    searchFor = ['Analytics Developers', 'CRIS Data Mart Support']
    df = df[~df['Assignment group'].str.contains('|'.join(searchFor), na=True)]

    #check if user selects milli/hd or both
    if config is 'milli':
        df = df[df['Configuration item'] == 'MILLI P134']
    if config is 'helpdesk':
        df = df[df['Configuration item'] != 'MILLI P134']
    if config is 'both':
        df = df
    else:
        print('Cannot create dataframe based on configuration')
    return(df)

#get esc/res stats per person
def personEscRes(df):
    df_opened = df.groupby('Opened by').size()
    df_closed = df.groupby('Opened by')['Assignment group'].apply(lambda x: x[x.str.contains('IS 4th Source Helpdesk')].count())
    df = pd.concat([df_closed, df_opened], axis=1)
    df.columns.values[1] = 'Opened'
    df.columns.values[0] = 'Resolved'
    df['Escalated'] = df['Opened'] - df['Resolved']
    df['Resolve Rate'] = df['Resolved']/df['Opened']
    print(df)

#get esc/res stats per Subcategory
def subEscRes(df):
    df_opened = df.groupby('Subcategory').size()
    df_closed = df.groupby('Subcategory')['Assignment group'].apply(lambda x: x[x.str.contains('IS 4th Source Helpdesk')].count())
    df = pd.concat([df_closed, df_opened], axis=1)
    df.columns.values[1] = 'Opened'
    df.columns.values[0] = 'Resolved'
    df['Escalated'] = df['Opened'] - df['Resolved']
    df['Resolve Rate'] = df['Resolved']/df['Opened']
    print(df)

def main():
    #create window object
    window = Tk()
    window.title('SJ Esc/Res Report Generator')

    def run(command):
        (str(command))
    
    def findConf():
        if is_checked1.get() and is_checked2.get():
            return('both')
        if is_checked1.get():
            return('helpdesk')
        if is_checked2.get():
            return('milli')

    def getFile():
        filename = StringVar()
        filename.set(tkFileDialog.askopenfilename(filetypes=(('SJ SNOW Excel Files', '.xlsx'), ('All files', '*.*'))))
        filePathEntry.insert(50, filename)

    def getPath():
        filePath = filePathEntry.get
        print(filePath)
        return(filePath)

    #define checkbuttons
    is_checked1 = IntVar()
    check1 = ttk.Checkbutton(window, text='Helpdesk', onvalue=1, offvalue=0, variable=is_checked1)
    check1.state(['!alternate'])
    check1.grid(row=3, column=0, sticky=W)
    is_checked2 = IntVar()
    check2 = ttk.Checkbutton(window, text='Milli', onvalue=1, offvalue=0, variable=is_checked2)
    check2.state(['!alternate'])
    check2.grid(row=4, column=0, sticky=W)

    #define buttons
    b1 = Button(window, text='Browse:', width=15, height=2, bg='blue', fg='white', command=lambda: run(getFile()))
    b1.grid(row=1, column=0)
    b1 = Button(window, text='Convert File to CSV:', width=15, height=2, bg='blue', fg='white', command=lambda: run(toCSV(getPath())))
    b1.grid(row=2, column=0)
    b1 = Button(window, text='Generate Report:', width=15, height=2, bg='orange', fg='black', command=lambda: run(personEscRes(createDF('report.csv', findConf()))))
    b1.grid(row=5, column=0)
    b1 = Button(window, text='Generate Report', width=15, height=2, bg='orange', fg='black', command=lambda: run(subEscRes(createDF('report.csv', findConf()))))
    b1.grid(row=6, column=0)
    b1 = Button(window, text='Get Entry', width=15, height=2, bg='orange', fg='black', command=getPath)
    b1.grid(row=7, column=0)

    #define labels
    L1 = ttk.Label(window, text='Menu')
    L1.grid(row=0,column=0)

    #define textboxes
    filePathEntry = Entry(window, width=50)
    filePathEntry.grid(row=1, column=1)
    
    window.mainloop()

main()
