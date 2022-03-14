# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
#           IMPORT                                                                #
# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
import os
import pandas as pd
import numpy as np
from __init__ import fix_text
import sqlite3
import pickle
import openpyxl as xl
import sys
_global = sys.modules[__name__] # Allows access to 'global' variables defined below
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
import time
import datetime
from author_diacritics import ensure_encryption
import re



# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
#           DISPLAY WINDOW                                                          #
# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #

def open_author_split():
    # Create window
    author_split = Tk()
    author_split.title("Author Split Program")
    author_split.configure(bg = '#003B49')
    author_split.iconbitmap('R:/storage/libarchive/b/1. Processing/8. Other Projects/Scholars-Mine-GitHub/Control Panel/CODE/S&T_Logo.ico')

    # Center the window on the screen
    window_w = 400 # window width
    window_h = 290 # window height

    screen_w = author_split.winfo_screenwidth() # screen width
    screen_h = author_split.winfo_screenheight() # screen height

    x_coor = (screen_w/2) - (window_w/2) #middle of screen x coordinate
    y_coor = (screen_h/2) #lower half of screen y coordinate

    author_split.geometry("%dx%d+%d+%d" % (window_w, window_h, x_coor, y_coor)) # place in middle

    #Create Progress Bar
    progress = ttk.Progressbar(author_split, orient = HORIZONTAL, length = 370, mode = 'determinate')

    # Create a frame
    frame = LabelFrame(author_split, text = "Select an Input File", bg = '#003B49', fg = 'white', font = 'tungsten 12 bold')

    # Create a label
    task = Label(author_split, text = "Waiting for a file", bg = '#003B49', fg = 'white', font = 'tungsten 12 bold')
    file_label = Label(author_split, text = "File: N/A", bg = '#003B49', fg = 'white', font = 'tungsten 12 bold')

    # Open Help Function
    def open_help():
        # Open word document
        try:
            os.startfile("R:/storage/libarchive/b/1. Processing/8. Other Projects/Scholars-Mine-GitHub/Control Panel/Documentation/Author Split Help.docx")
        except:
            error_popup("Could not open help file.")

    # Create a button
    help_button = Button(author_split, text = "Help", command = lambda : open_help(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    exit_button = Button(author_split, text = "Exit Author Split", command = author_split.destroy, bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")

    # Place in window
    frame.pack(padx = 20, pady = 10)
    file_label.pack(padx = 20, pady = (0, 20))
    progress.pack(padx = 5, pady = 5)
    task.pack(padx = 20)
    help_button.pack(padx = (331, 10), pady = 0)
    exit_button.pack(padx = (242, 10), pady = (5, 0))

    # Error Message Popup
    def error_popup(error_message):
        messagebox.showerror("Error", error_message)

    # Warning Message Popup
    def warning_popup(warning_message):
        messagebox.showwarning("Warning", warning_message)

    # Update Progress Bar
    def update_progress(p, t):
        # Update bar value
        progress['value'] = (p/2)*100

        # Update bar label
        task.config(text = t)
        task.pack()

        # Refresh window (very important line!)
        author_split.update_idletasks()
        time.sleep(0.2)

    # Function Click
    def browse():
        # Reset progress bar
        update_progress(0, "Waiting for a file")

        # Open file
        author_split.filename = filedialog.askopenfilename(initialdir = "R:/storage/libarchive/b/1. Processing/8. Other Projects/Excel Files", title = "Select Input", filetypes = (("Excel Workbook", "*.xlsx"),))
        if author_split.filename == "":
            warning_popup("No file selected.")
            file_label.config(text = "File: N/A")
            del author_split.filename
        else:
            #Get file name
            name = author_split.filename
            for i in range(len(name)):
                if "/" in name[-(i)]:
                    name = "File: " + name[-(i-1):]
                    break

            #Update Folder Name Label
            file_label.config(text = name)

    # Main Function
    def author_split_main(file_name):
        update_progress(1, "Splitting Authors...")
                
        rdsheet = None
        author_column = ''
        excelName = ''

        authority_database = sqlite3.connect('R:/storage/libarchive/b/1. Processing/8. Other Projects/Scholars-Mine-GitHub/Control Panel/CODE/faculty-author-split-test.db')
        authority_cursor = authority_database.cursor()
        def regexp(expr, item):
            reg = re.compile(expr)
            return reg.search(item) is not None
        authority_database.create_function("REGEXP", 2, regexp)

        sqlite3.enable_callback_tracebacks(True)   # <-- !\

        authorDict = {}
        rb = None

        # Read in Excel file
        with pd.ExcelFile(file_name) as excel_in:
            df = pd.read_excel( excel_in, excel_in.sheet_names[0], header=0, index_col=False )

        for i, c in enumerate( df.columns ):
            if "_fname" in c:
                df.iloc[ :, i ] = df.iloc[ :, i ].fillna('')
                df.iloc[ :, i + 1 ] = df.iloc[ :, i + 1 ].fillna('')
                for j in range( df.shape[0] ):
                    split_name = df.iloc[ j, i ].split(" ")
                    df.iloc[ j, i ] = split_name[0]
                    split_name = split_name[1:]
                    df.iloc[ j, i + 1 ] = " ".join( split_name )

        # Find columns corresponding to S&T affiliates, that is, columns whose
        ## "autho<i>_institution" is "Missouri University of Science and Technology"
        author_columns = []
        for r in range( df.shape[0] ):
            for i, c in enumerate(df.columns):
                if "_institution" in c and df.loc[r, c] == "Missouri University of Science and Technology":
                    author_columns.append( (r, i-5) )

        for c in df.columns:
            if "_institution" in c:
                # print(c, i+1)
                t_df = df.loc[:, c] = ""

        # $author_columns contains S&T graduate students, to filter them out, we will
        ## filter them out by looking them up in the faculty.db
        df["faculty_author_count"] = np.zeros( df.shape[0] )
        dct_df = {}
        for index in author_columns:
            t_df = df.iloc[ index[0], index[1]:index[1]+7 ]

            t_name_dct = { "first": t_df.values[0].split(" ")[0]
                            , "last":t_df.values[2].split(" ")[0]
                        }

            t_query = authority_cursor.execute('SELECT authority_name,first_name, middle_name, last_name, email, department FROM faculty WHERE (last_name COLLATE NOCASE) LIKE :last AND (first_name COLLATE NOCASE) LIKE :first', t_name_dct).fetchall()
            if len(t_query) == 1:
                dct_df[ index ] = list(t_query[0])
                df.loc[ index[0], "faculty_author_count" ] += 1
            elif len(t_query) > 1:
                t_query = list(t_query[0])
                for i in [0,2,4,5]:
                    t_query[i] = ""
                t_query[-1] = "MANUAL-CHECK"
                dct_df[ index ] = t_query
                df.loc[ index[0], "faculty_author_count" ] = np.nan
            else:
                for key in t_name_dct:
                    t_name_dct[key] = t_name_dct[key].replace(".","").replace(",","")
                t_query = authority_cursor.execute('SELECT authority_name, first_name, middle_name, last_name, email, department FROM faculty WHERE (last_name COLLATE NOCASE) LIKE :last AND first_name REGEXP :first', t_name_dct).fetchall()
                if len(t_query) == 1:
                    dct_df[ index ] = list(t_query[0])
                    df.loc[ index[0], "faculty_author_count" ] += 1
                elif len(t_query) > 1:
                    t_query = list(t_query[0])
                    for i in [0,2,4,5]:
                        t_query[i] = ""
                    t_query[-1] = "MANUAL-CHECK"
                    dct_df[ index ] = t_query
                    df.loc[ index[0], "faculty_author_count" ] = np.nan

        '''
            For each S&T author, fill in
                author{i}_fname	author{i}_mname	author{i}_lname	author{i}_suffix	author{i}_email	author{i}_institution	author{i}_is_corporate
            with data taken from the faculty database and setting the institution to "Missouri University of Science and Technology"
        '''
        for index in dct_df:
            try:
                df.iloc[ index[0], index[1]:index[1]+7 ] = dct_df[index][1:4] + [""] + dct_df[index][4:5] + ["Missouri University of Science and Technology",""]
            except:
                print(dct_df[index], end="\n\n\n")

        '''
            Get total author count by counting non-empty last name columns
        '''
        df = df.fillna('')
        df["total_author_count"] = np.zeros(df.shape[0])
        for c in df.columns:
            for r in range( df.shape[0] ):
                if "_lname" == c[-6:] and df.loc[ r, c ]:
                    # print(str(df.loc[ r, c ]) )
                    df.loc[r, "total_author_count"] += 1

        '''
            (1) Get authorized names separated by '<br>'
            (2) Get departments 1-4, non-repeating, in order of authors
        '''
        df["authorized_name"] = ""
        for i in range(4):
            df[ f"department{i+1}" ] = ""
        for r in range( df.shape[0] ):
            department_num = 1
            authorized_name_list = ""
            for index in dct_df:
                if index[0] == r:
                    authorized_name_list += dct_df[index][0] + "<br>"
                    if department_num <= 4:
                        department = dct_df[index][-1]
                        department_runner = np.max( [1, department_num-1] )
                        while department_runner >= 1:
                            if department == df.loc[r, f"department{department_runner}"]:
                                break
                            department_runner -= 1
                        # print( department_runner )
                        if department_runner == 0:
                            # print(department)
                            df.loc[r, f"department{department_num}"] = department
                            department_num += 1
            authorized_name_list = authorized_name_list[:-4]
            df.loc[ r, "authorized_name" ] = authorized_name_list

        '''
            Replace diacritics with appropriate token in columns: Abstract, Keywords, Funding Sponsor, Publisher, First name, Middle name, Last names, Source Publications
        '''
        df = ensure_encryption(df)

        # Write in Excel file
        with pd.ExcelWriter(f"{file_name[:-5]}_Author_Split.xlsx") as excel_out:
            df.to_excel(excel_out, sheet_name="Sheet1", header=True, index=False)
            #excel_out.close()
            del df

        update_progress(2, "Excel Created")

    # Start Button Function
    def start():
        # Run program and update progress bar
        #try:
        author_split_main(author_split.filename)
        #except AttributeError:
            #error_popup("No folder selected. Browse to select a folder.")
        #except:
        #    error_popup("There was an unknown error, the file could not be processed.")

    # Create a button
    browse_button = Button(frame, text = "Browse", command = lambda : browse(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    start_button = Button(frame, text = "Start", command = lambda : start(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")

    # Place in frame
    browse_button.grid(row = 0, column = 0, padx = (15, 0), pady = 15)
    start_button.grid(row = 0, column = 1, padx = 15, pady = 15)

    # Keeps window open until closed
    author_split.mainloop()
