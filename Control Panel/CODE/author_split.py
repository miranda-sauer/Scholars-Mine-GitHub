# cd /d "R:\storage\libarchive\b\1. Processing\8. Other Projects\Scholars-Mine-GitHub\Stand Alone Author Split" 
# python author_.py
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


# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
#           GLOBAL VARIABLES                                                        #
# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #

rdsheet = None
author_column = ''
excelName = ''

authority_database = sqlite3.connect('R:/storage/libarchive/b/1. Processing/8. Other Projects/Scholars-Mine-GitHub/Stand Alone Author Split/faculty-author-split-test.db')
authority_cursor = authority_database.cursor()
def regexp(expr, item):
    reg = re.compile(expr)
    return reg.search(item) is not None
authority_database.create_function("REGEXP", 2, regexp)

# print(authority_cursor.execute('SELECT authority_name,first_name, middle_name, last_name, email, department FROM faculty WHERE (last_name COLLATE NOCASE) LIKE "Chen" AND first_name REGEXP "G"').fetchall())
# quit()

authorDict = {}
rb = None

# special_char = pickle.load(open('R:/storage/libarchive/b/1. Processing/8. Other Projects/Scholars-Mine-GitHub/Stand Alone Author Split/special_char.pickle','rb'))
# extra_special_char = pickle.load(open('R:/storage/libarchive/b/1. Processing/8. Other Projects/Scholars-Mine-GitHub/Stand Alone Author Split/extra_special_char.pickle','rb'))

# Read in Excel file
excel_file = r"R:\storage\libarchive\b\1. Processing\8. Other Projects\Excel Files\Epting-new_format_harvest.xlsx"
with pd.ExcelFile( excel_file ) as excel_in:
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
            # # print(c, i+1)
            # t_df = df.loc[df[c] == "Missouri University of Science and Technology"]
            # # print(t_df)
            # if t_df.shape[0] > 0:
            #     for r in t_df.index:
            #         author_columns.append( (r, i-5) )

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
    # t_query = authority_cursor.execute(f'SELECT authority_name,first_name, middle_name, last_name, email FROM faculty WHERE (name_variations COLLATE NOCASE) LIKE :lookup', t_name_dct).fetchall()
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
        # print(t_name_dct)
        t_query = authority_cursor.execute('SELECT authority_name,first_name, middle_name, last_name, email, department FROM faculty WHERE (last_name COLLATE NOCASE) LIKE :last AND first_name REGEXP :first', t_name_dct).fetchall()
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
with pd.ExcelWriter(f"{excel_file[:-5]}_Completed.xlsx") as excel_out:
    df.to_excel(excel_out, sheet_name="Sheet1", header=True, index=False)
    #excel_out.close()
    del df

# df.to_csv("test.csv", header=True, index=False, encoding='utf-8')
