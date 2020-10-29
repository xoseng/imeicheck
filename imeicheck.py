# coding=utf8
# author: Xose Brais Noya Garcia
# IMPORTS

import tkinter_library_functions
import cmd_command

from PIL import Image
from tkinter import *
from tkinter import filedialog
from PIL import ImageTk, Image
import webbrowser
import threading
from tkinter import ttk
import pandas as pd
import openpyxl
import lxml
import xlsxwriter

def main_start():

    # NO TK FUNCTIONS
    def clear_vals():
        val_filelist.set('')
        val_workfile.set('')
        val_numfile.set(0)
        val_check.set(0)
        val_check0.set(0)
        val_check2.set(0)
        val_check3.set(0)
        try:
            if val_filepath.get() != '':
                cmd_command.wincmd('del "' + val_filepath.get() + '"')
        except:
            pass #dont exist
        val_filepath.set('')

    def set_file():
        pathtoset = tkinter_library_functions.get_filepathload()
        val_filelist.set(pathtoset)
        num_files=len(pathtoset)
        val_numfile.set(num_files)
        if num_files == 0:
            num_files=str(num_files)
            val_workfile.set('No files selected!')
        elif num_files == 1:
            val_workfile.set('You have selected ' + str(num_files) + ' file!')
        else:
            val_workfile.set('You have selected '+str(num_files)+' files!')
    def start_work():
        if val_filelist.get() == '':
            clear_vals()
            tkinter_library_functions.warning_file()
        else:
            filepath=tkinter_library_functions.set_filepathsave()
            val_filepath.set(filepath)
            if filepath == '':
                clear_vals()
                tkinter_library_functions.warning_path()
            else:
                if val_check0.get() == 0 and val_check.get() == 0 and val_check2.get() == 0 and val_check3.get() == 0:
                    clear_vals()
                    tkinter_library_functions.warning_option()
                else:
                    filelist=val_filelist.get()
                    files_number=val_numfile.get()
                    files_number=files_number
                    i=0
                    while (i < files_number):
                        file_path=filelist.split(',')[i]
                        file_path=file_path.split("'")[1]
                        file_path=file_path.replace("/", "\\")
                        file_path=file_path.strip()
                        #crear documento temporal para meter los valores
                        df = pd.read_excel(file_path)
                        csvfile= 'temp' + str(i) + '.csv'
                        df.to_csv(csvfile, index=False)
                        i=i+1

                    csvcomplete_name='temp.csv'
                    with open(csvcomplete_name, "w") as csvcomplete:
                        i = 0
                        while (i < files_number):
                            csvfile = 'temp' + str(i) + '.csv'
                            with open(csvfile, "r") as list:
                                for line in list:
                                    if ',' in line:
                                        line = line.replace(',', '\n')
                                    csvcomplete.write(line)
                            i=i+1
                            cmd_command.wincmd('del "'+csvfile+'"')

                    list_final = []
                    with open(csvcomplete_name, "r") as listfinal:
                        for line in listfinal:
                            line = line.replace('\n', '')
                            if any(line):
                                if '.' in line:
                                    line = line.split('.')[0]
                                list_final.append(line)
                    cmd_command.wincmd('del "' + csvcomplete_name + '"')
                    repeated = []
                    single = []
                    notimei = []

                    for x in list_final:
                        long = len(x)
                        num = x.isnumeric()
                        #update to set list with no IMEIS
                        if long != 15 or num != True:
                            notimei.append(x)
                        #end update
                        if x not in single:
                            if long == 15 and num == True:
                                single.append(x)
                        else:
                            if x not in repeated:
                                if long == 15 and num == True:
                                    repeated.append(x)

                        single_csv='single.csv'
                        repeated_csv='repeated.csv'
                        uniques_csv='uniques.csv'
                        notimei_csv = 'notimei.csv'

                        if val_check0.get() != 0:
                            with open(single_csv, 'w') as singlecsv:
                                singlecsv.write('IMEI FULL LIST\n')
                                for element in single:
                                    singlecsv.write(element+'\n')
                        if val_check.get() != 0:
                            with open(repeated_csv, 'w') as repeatedcsv:
                                repeatedcsv.write('REPEATED FOUND\n')
                                for element in repeated:
                                    repeatedcsv.write(element+'\n')
                        if val_check3.get() != 0:
                            with open(notimei_csv, 'w') as notimeicsv:
                                notimeicsv.write('NOT IMEIS\n')
                                for element in notimei:
                                    notimeicsv.write(element+'\n')

                    uniques = []
                    for z in single:
                        if z not in repeated:
                            uniques.append(z)
                    if val_check2.get() != 0:
                        with open(uniques_csv, 'w') as uniquescsv:
                            uniquescsv.write('NOT REPEATED\n')
                            for element in uniques:
                                uniquescsv.write(element + '\n')

                    single_xlsx='single.xlsx'
                    repeated_xlsx='repeated.xlsx'
                    uniques_xlsx='uniques.xlsx'
                    notimei_xlsx = 'notimei.xlsx'

                    if val_check0.get() != 0:
                        excel_namepage = 'IMEIS'
                        df = pd.read_csv(single_csv).astype(str)
                        df.to_excel(single_xlsx, excel_namepage, index=None, header=True)
                    if val_check.get() != 0:
                        excel_namepage = 'REPEATED FOUND'
                        df = pd.read_csv(repeated_csv).astype(str)
                        df.to_excel(repeated_xlsx, excel_namepage, index=None, header=True)
                    if val_check2.get() != 0:
                        excel_namepage = 'NOT REPEATED'
                        df = pd.read_csv(uniques_csv).astype(str)
                        df.to_excel(uniques_xlsx, excel_namepage, index=None, header=True)
                    if val_check3.get() != 0:
                        excel_namepage = 'NOT IMEIS'
                        df = pd.read_csv(notimei_csv).astype(str)
                        df.to_excel(notimei_xlsx, excel_namepage, index=None, header=True)

                    try:
                        cmd_command.wincmd('del "' + single_csv + '"')
                    except:
                        pass
                    try:
                        cmd_command.wincmd('del "' + repeated_csv + '"')
                    except:
                        pass
                    try:
                        cmd_command.wincmd('del "' + uniques_csv + '"')
                    except:
                        pass
                    try:
                        cmd_command.wincmd('del "' + notimei_csv + '"')
                    except:
                        pass

                    final_xlsx='final.xlsx'

                    writer = pd.ExcelWriter(final_xlsx, engine='xlsxwriter')

                    if val_check0.get() != 0:
                        xls_file = pd.ExcelFile(single_xlsx)
                    if val_check.get() != 0:
                        xls_file2 = pd.ExcelFile(repeated_xlsx)
                    if val_check2.get() != 0:
                        xls_file3 = pd.ExcelFile(uniques_xlsx)
                    if val_check3.get() != 0:
                        xls_file4 = pd.ExcelFile(notimei_xlsx)

                    if val_check0.get() != 0:
                        df = xls_file.parse('IMEIS')
                        df.to_excel(writer, sheet_name='IMEIS', startrow=1, header=False, index=False)

                    if val_check.get() != 0:
                        df2 = xls_file2.parse('REPEATED FOUND')
                        df2.to_excel(writer, sheet_name='REPEATED FOUND', startrow=1, header=False, index=False)

                    if val_check2.get() != 0:
                        df3 = xls_file3.parse('NOT REPEATED')
                        df3.to_excel(writer, sheet_name='NOT REPEATED', startrow=1, header=False, index=False)

                    if val_check3.get() != 0:
                        df4 = xls_file4.parse('NOT IMEIS')
                        df4.to_excel(writer, sheet_name='NOT IMEIS', startrow=1, header=False, index=False)

                    workbook = writer.book

                    header_format = workbook.add_format({
                        'bold': True,
                        'fg_color': '#008000',
                        'color': 'white',
                        'align': 'center',
                        'border': 1})
                    header_format2 = workbook.add_format({
                        'bold': True,
                        'fg_color': '#FF0000',
                        'color': 'white',
                        'align': 'center',
                        'border': 1})
                    header_format3 = workbook.add_format({
                        'bold': True,
                        'fg_color': '#008080',
                        'color': 'white',
                        'align': 'center',
                        'border': 1})
                    header_format4 = workbook.add_format({
                        'bold': True,
                        'fg_color': '#800080',
                        'color': 'white',
                        'align': 'center',
                        'border': 1})
                    format_text = workbook.add_format({'num_format': '0', 'align': 'center'})

                    if val_check0.get() != 0:
                        worksheet = writer.sheets['IMEIS']
                        for col_num, value in enumerate(df.columns.values):
                            worksheet.write(0, col_num, value, header_format)
                            column_len = df[value].astype(str).str.len().max()
                            # Setting the length if the column header is larger
                            # than the max column value length
                            column_len = max(column_len, len(value)) + 3
                            # print(column_len)
                            # set the column length
                            # worksheet.set_column(col_num, col_num, column_len)
                            worksheet.set_column(col_num, col_num, column_len, format_text)

                    if val_check.get() != 0:
                        worksheet = writer.sheets['REPEATED FOUND']
                        for col_num, value in enumerate(df2.columns.values):
                            worksheet.write(0, col_num, value, header_format2)
                            column_len = df2[value].astype(str).str.len().max()
                            # Setting the length if the column header is larger
                            # than the max column value length
                            column_len = max(column_len, len(value)) + 3
                            # print(column_len)
                            # set the column length
                            # worksheet.set_column(col_num, col_num, column_len)
                            worksheet.set_column(col_num, col_num, column_len, format_text)

                    if val_check2.get() != 0:
                        worksheet = writer.sheets['NOT REPEATED']
                        for col_num, value in enumerate(df3.columns.values):
                            worksheet.write(0, col_num, value, header_format3)
                            column_len = df3[value].astype(str).str.len().max()
                            column_len = max(column_len, len(value)) + 3
                            worksheet.set_column(col_num, col_num, column_len, format_text)

                    if val_check3.get() != 0:
                        worksheet = writer.sheets['NOT IMEIS']
                        for col_num, value in enumerate(df4.columns.values):
                            worksheet.write(0, col_num, value, header_format4)
                            column_len = df4[value].astype(str).str.len().max()
                            column_len = max(column_len, len(value)) + 3
                            worksheet.set_column(col_num, col_num, column_len, format_text)

                    writer.save()

                    clear_vals()

                    try:
                        cmd_command.wincmd('del "' + single_xlsx + '"')
                    except:
                        pass
                    try:
                        cmd_command.wincmd('del "' + repeated_xlsx + '"')
                    except:
                        pass
                    try:
                        cmd_command.wincmd('del "' + uniques_xlsx + '"')
                    except:
                        pass
                    try:
                        cmd_command.wincmd('del "' + notimei_xlsx + '"')
                    except:
                        pass
                    try:
                        cmd_command.wincmd('del "' + filepath + '"')
                    except:
                        pass
                    try:
                        cmd_command.wincmd('copy "' + final_xlsx + '" "' + filepath + '"')
                    except:
                        pass
                    try:
                        cmd_command.wincmd('del "' + final_xlsx + '"')
                    except:
                        pass
                    try:
                        cmd_command.wincmd('"' +filepath+'"')
                    except:
                        pass

    #MAIN
    version_software = 'IMEI Check RC V2.3'
    logo_path='./img/logo.ico'
    root = Tk()
    root.config(bd=30)
    root.title(version_software)  # Titulo de la ventana
    root.iconbitmap(logo_path)  # Icono de la ventana, en ico o xbm en Linux
    root.resizable(0, 0)  # Desactivar redimension de ventana

    menubar = Menu(root)
    root.config(menu=menubar)

    helpmenu = Menu(menubar, tearoff=0)
    helpmenu.add_command(label="License agreement", command=tkinter_library_functions.license_agreement_gnu)
    helpmenu.add_command(label="Project website", command=tkinter_library_functions.project_website)
    helpmenu.add_command(label="About", command=tkinter_library_functions.about_program)
    helpmenu.add_command(label="Exit", command=tkinter_library_functions.exit_program)
    menubar.add_cascade(label="Info", menu=helpmenu)

    imgpath = './img/logo.png'
    img_tk = ImageTk.PhotoImage(Image.open(imgpath))
    imglabel = Label(root, image=img_tk).grid(row=1, column=1, padx=5, pady=5)

    ######################GLOBAL TK VARS DECLARATION

    val_workfile = StringVar()
    val_filelist = StringVar()
    val_check0 = IntVar(value=0)
    val_check = IntVar(value=0)
    val_check2 = IntVar(value=0)
    val_check3 = IntVar(value=0)
    val_numfile=IntVar()
    val_filepath = StringVar()

    #BUTTONS LABEL FORMS...

    workfile_button=Button(root, justify="left", text="SELECT FILES", command=set_file).grid(row=2, column=0, padx=5, pady=2)
    workfile_value=Entry(root, justify="center", textvariable=val_workfile, state="disabled", width=40).grid(row=2, column=1, padx=5, pady=2)
    reset_button=Button(root, justify="left", text="RESTART", command=clear_vals).grid(row=2, column=2, padx=5, pady=2)
    checkduplicated_button = Checkbutton(root, justify="left", text="SHOW SHEET WITH ALL UNIQUE IMEIS    ", variable=val_check0, onvalue=1, offvalue=0).grid(row=3, column=1, padx=5, pady=2)
    checkduplicated_button=Checkbutton(root, justify="left", text="SHOW SHEET WITH DUPLICATED IMEIS    ", variable=val_check, onvalue=1, offvalue=0).grid(row=4, column=1, padx=5, pady=2) #set whitespace to simple align
    checkunique_button = Checkbutton(root, justify="left", text="SHOW SHEET WITH NOT REPEATED IMEIS", variable=val_check2, onvalue=1, offvalue=0).grid(row=5, column=1, padx=5, pady=2)
    checknotimei_button = Checkbutton(root, justify="left", text="SHOW SHEET WITH NOT IMEIS VALUES    ", variable=val_check3, onvalue=1, offvalue=0).grid(row=6, column=1, padx=5, pady=2)
    starbutton=Button(root, justify="left", text="START", command=start_work).grid(row=7, column=1, padx=5, pady=5)

    # CENTER WINDOW TO SCREEN
    tkinter_library_functions.center(root)
    # LOOP TK
    root.mainloop()

main_start()
