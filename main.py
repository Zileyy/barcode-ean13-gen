#IMPORTS
from tkinter import filedialog
import pandas as pd
import tkinter as tk
from barcode import EAN13
from barcode.writer import ImageWriter
import os


root = tk.Tk()      #Tkinter root
xlsxPath = ''       #Path of an excel file

#FUNCTIONS
#Function that imports excel file with data
def importPath():
    xlsxPath = filedialog.askopenfilename()
    print(str(xlsxPath))
    data = pd.read_excel(str(xlsxPath))
    return data

#VARS
data = importPath()                 #Data from excel file
barcodes = data['Barcode'].tolist() #Barcodes
names = data['Name'].tolist()       #Names from excel sheet
edited_names =[]                    #Names from excel sheet with removed spaces and converted to lowercase
values = {}                         #Barcodes stored with the names

#Function that converts names to all lowercase without spaces for export
def convertNames(names):
    for x in names:
        x = x.replace(' ', '')
        x = x.lower()
        edited_names.append(x)

#Function that stores name and barcode into dict
def storeValues(n,b):
    values = zip(n,b)
    values = dict(values)
    return values

#Function that generates Barcodes
def generate(vars,exprt):
    perma = []
    #Checks for every name and barcode associated with it in dict
    for key, value in vars.items() :
        #var export combines location of the files and file names
        export = str(exprt+'/'+key+'.jpg')
        perma.append(str(export))
        #Exports Barcode with the name from the excel sheet to the same location of excel file
        EAN13(str(value), writer=ImageWriter()).write(export)
    f = open('paths.csv','w')
    for x in perma:
        f.write(x)
        f.write('\n')
    f.close()

#Function that runs all modules needed for task
def run(exprt):
    convertNames(names)
    values = storeValues(edited_names,barcodes)
    generate(values,exprt)

#Function that props windows "open" option
def gen():
    export_path = tk.filedialog.askdirectory()
    run(export_path)

#TKINTER UI

root.minsize('350','320')
root.maxsize('350','320')

main_frame = tk.Frame(root)
main_frame.pack(side='top')

title_panel = tk.Label(main_frame, text='Excel Barcode Generator')
title_panel.configure(bg='lightblue', fg='black', width=45)
title_panel.pack(side='top')

break_panel = tk.Label(main_frame)
break_panel.pack(side='top')

lower_frame = tk.Frame(root)
lower_frame.pack(side='top')


generate_text = tk.Label(lower_frame, text='Generate Barcodes')
generate_text.pack(side='top')

btnGenerate = tk.Button(lower_frame, text='Generate', fg='black', command=gen)
btnGenerate.pack(side='top')

break_panel2 = tk.Label(lower_frame)
break_panel2.pack(side='top')


bottom_frame = tk.Frame(root)
bottom_frame.pack(side='bottom')

status_info_text = tk.Label(bottom_frame, text='status: Note file must have columns Barcode and Name')
status_info_text.pack(side='top')

end_panel = tk.Label(bottom_frame, text='')
end_panel.configure(bg='lightblue', fg='black', width=45)
end_panel.pack(side='top')

root.mainloop()

os.system('pause')
