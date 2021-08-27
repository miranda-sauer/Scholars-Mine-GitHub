# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
#           IMPORT                                                                #
# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
import datetime
import sqlite3
import time
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter
from openpyxl.styles import Alignment, PatternFill
from openpyxl.styles.colors import Color
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
import os



# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
#           REMOVE DASH                                                             #
# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #

def remove_dash(col_value):
    index = 0                                                                       # Create index

    while index >= 0 and index < len(col_value):                                    # Search through string with a valid index
        if '-' in col_value[index]:                                                 # Find dash
            col_value = col_value[:index] + ' thru ' + col_value[index + 1:]        # Remove dash and replace with thru
        index += 1                                                                  # Increment index

    return col_value                                                                # Return string without dash



# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
#           CHANGE DATE TO NUM                                                      #
# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #

def date_to_num(issnum):
    # Get parts from date
    month = issnum[5:7]
    day = issnum[8:10]

    # Get rid of 0s
    if '0' in month[0]:
        month = month[1:]

    if '0' in day[0]:
        day = day[1:]

    # Format output
    if int(day) == int(month):
        output = day
    elif int(day) > int(month):
        output = month + ' thru ' + day
    elif int(day) < int(month):
        output = day + ' thru ' + month
    return output



# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
#           CAPITALIZE LETTER                                                       #
# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #

def manual_upper(title):
    fixed_title = ''    #value returned
    words = []          #list of words in title
    index = []          #list of indicies for uppercase words
    count = 0           #index for uppercase words
    start = 0           #used for finding words in title
    stop = 1            #used for finding words in title

    # Find words in title and add them to words[]
    while start < len(title) and stop < len(title) and start < stop:
        if ' ' in title[stop]:
            words.append(str(title[start:stop]) + ' ')
            start = stop + 1
            stop += 1
        if stop == len(title) - 1:
            words.append(str(title[start:stop + 1]))
        stop += 1

    # Finds uppercase words and adds index to index[]
    for element in words:
        if element.isupper() == True:
            index.append(count)
        count += 1

    # Depending on percentage of uppercase words, makes words lowercase
    if (len(index)/len(words)*100) > 50: #percentage of uppercase words
        for value in index:
            words[value] = (words[value].lower())

    # Puts all fixed words back into one string
    for element in words:
        fixed_title = fixed_title + element

    #Capitalizes first letter in title
    fixed_title = fixed_title[0].upper() + fixed_title[1:]

    return fixed_title



# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
#           FIX SUBSCRIPT AND SUPERSCRIPT                                           #
# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #

def replace_sub_sup(title):
    while('</sub>' in title):                           #if there is a sub to replace...
        index = 0                                           #start at the beginning of the string
        while(index < len(title)):                         #while the index is in the string
            if '<sub>' in title[index:index + 5]:               #find were '<sub>' is in the string
                title = title[0:index] + title[index + 5:]          #get rid of '<sub>' in the string
                while(title[index] != '<'):                         #while the index isn't '<'
                    title = str(title[0:index]) + str(char_to_sub(title[index])) + str(title[index + 1:])  #change char into subscript
                    index += 1                                                              #increment the index
                title = title[0:index] + title[index + 6:]          #get rid of '</sub>' in the string
            index += 1                                          #increment the index

    while('</sup>' in title):                           #if there is a sup to replace...
        index = 0                                           #start at the beginning of the string
        while(index < len(title)):                         #while the index is in the string
            if '<sup>' in title[index:index + 5]:               #find were '<sup>' is in the string
                title = title[0:index] + title[index + 5:]          #get rid of '<sup>' in the string
                while(title[index] != '<'):                         #while the index isn't '<'
                    title = str(title[0:index]) + str(char_to_sup(title[index])) + str(title[index + 1:])  #change char into subscript
                    index += 1                                                              #increment the index
                title = title[0:index] + title[index + 6:]          #get rid of '</sup>' in the string
            index += 1                                          #increment the index
    return title



# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
#           CHARACTER TO SUBSCRIPT                                                  #
# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #

def char_to_sub(argument):
    output = ''
    switcher = {
        '0': '\u2080',
        '1': '\u2081',
        '2': '\u2082',
        '3': '\u2083',
        '4': '\u2084',
        '5': '\u2085',
        '6': '\u2086',
        '7': '\u2087',
        '8': '\u2088',
        '9': '\u2089',
        '+': '\u208A',
        '-': '\u208B',
        '=': '\u208C',
        '(': '\u208D',
        ')': '\u208E',
        'a': '\u2090',
        'e': '\u2091',
        'o': '\u2092',
        'x': '\u2093',
        'h': '\u2095',
        'k': '\u2096',
        'l': '\u2097',
        'm': '\u2098',
        'n': '\u2099',
        'p': '\u209A',
        's': '\u209B',
        't': '\u209C',
    }
    output = switcher.get(argument, '?')
    if '?' in output:
        SUB = str.maketrans("aehijklmnoprstuvxy", "ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓᵧ")
        output = argument.translate(SUB)
    return output



# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
#           CHARACTER TO SUPERSCRIPT                                                #
# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #

def char_to_sup(argument):
    output = ''
    switcher = {
        '0': '\u2070',
        '1': '\u00B9',
        '2': '\u00B2',
        '3': '\u00B3',
        '4': '\u2074',
        '5': '\u2075',
        '6': '\u2076',
        '7': '\u2077',
        '8': '\u2078',
        '9': '\u2079',
        'i': '\u2071',
        '+': '\u207A',
        '-': '\u207B',
        '=': '\u207C',
        '(': '\u207D',
        ')': '\u207E',
        'n': '\u207F',
    }
    output = switcher.get(argument, '?')
    if '?' in output:
        SUP = str.maketrans("ABDEGHIJKLMNOPRTUVWabcdefghijklmnoprstuvwxyz", "ᴬᴮᴰᴱᴳᴴᴵᴶᴷᴸᴹᴺᴼᴾᴿᵀᵁⱽᵂᵃᵇᶜᵈᵉᶠᵍʰⁱʲᵏˡᵐⁿᵒᵖʳˢᵗᵘᵛʷˣʸᶻ")
        output = argument.translate(SUP)
    return output



# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
#           SEARCH DATABASE FOR AUTHORS                                             #
# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #

def search_database(l_name, f_name):
    conn = sqlite3.connect('R:/storage/libarchive/b/1. Processing/8. Other Projects/Scholars-Mine-GitHub/Control Panel/CODE/faculty.db')

    c = conn.cursor()

    print("Looking for... " + str(l_name) + ', ' + str(f_name))
    print()

    c.execute("SELECT * FROM faculty WHERE last_name = ? AND first_name = ?", (l_name,f_name))

    people = c.fetchall()

    for person in people:
        for num in range(0, 10):
            print(person[num])
        print('* * * * * * * * * * * * * * * * * * * * * * * * * * * * *')
        if person[7]:
                conn.close()
                return person[7]

    conn.close()



# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
#           DISPLAY WINDOW                                                          #
# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #

def open_harvesting():
    # Create window
    harvesting = Toplevel()
    harvesting.title("Harvesting Program")
    harvesting.configure(bg = '#003B49')
    harvesting.iconbitmap('R:/storage/libarchive/b/1. Processing/8. Other Projects/Scholars-Mine-GitHub/Control Panel/CODE/S&T_Logo.ico')

    # Center the window on the screen
    window_w = 400 # window width
    window_h = 290 # window height

    screen_w = harvesting.winfo_screenwidth() # screen width
    screen_h = harvesting.winfo_screenheight() # screen height

    x_coor = (screen_w/2) - (window_w/2) #middle of screen x coordinate
    y_coor = (screen_h/2) #lower half of screen y coordinate

    harvesting.geometry("%dx%d+%d+%d" % (window_w, window_h, x_coor, y_coor)) # place in middle

    #Create Progress Bar
    progress = ttk.Progressbar(harvesting, orient = HORIZONTAL, length = 370, mode = 'determinate')

    # Create a frame
    frame = LabelFrame(harvesting, text = "Select an Input File", bg = '#003B49', fg = 'white', font = 'tungsten 12 bold')

    # Create a label
    task = Label(harvesting, text = "Waiting for a file", bg = '#003B49', fg = 'white', font = 'tungsten 12 bold')
    file_label = Label(harvesting, text = "File: N/A", bg = '#003B49', fg = 'white', font = 'tungsten 12 bold')

    # Open Help Function
    def open_help():
        # Open word document
        try:
            os.startfile("**/*/Control Panel/Documentation/Harvesting Help.docx")
        except:
            error_popup("Could not find a help file to open.")


    # Create a button
    help_button = Button(harvesting, text = "Help", command = lambda : open_help(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    exit_button = Button(harvesting, text = "Exit Harvesting", command = harvesting.destroy, bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")

    # Place in window
    frame.pack(padx = 20, pady = 10)
    file_label.pack(padx = 20, pady = (0, 20))
    progress.pack(padx = 5, pady = 5)
    task.pack(padx = 20)
    help_button.pack(padx = (331, 10), pady = 0)
    exit_button.pack(padx = (250, 10), pady = (5, 0))

    # Error Message Popup
    def error_popup(error_message):
        messagebox.showerror("Error", error_message)

    # Warning Message Popup
    def warning_popup(warning_message):
        messagebox.showwarning("Warning", warning_message)

    # Update Progress Bar
    def update_progress(p, t):
        # Update bar value
        progress['value'] = (p/9)*100

        # Update bar label
        task.config(text = t)
        task.pack()

        # Refresh window (very important line!)
        harvesting.update_idletasks()
        time.sleep(0.2)

    # Function Click
    def browse():
        # Reset progress bar
        update_progress(0, "Waiting for a file")

        # Open file
        harvesting.filename = filedialog.askopenfilename(initialdir = "R:/storage/libarchive/b/1. Processing/8. Other Projects/Excel Files", title = "Select Input", filetypes = (("Excel Workbook", "*.xlsx"),))
        if harvesting.filename == "":
            warning_popup("No file selected.")
            file_label.config(text = "File: N/A")
            del harvesting.filename
        else:
            # Get file name
            name = harvesting.filename
            for i in range(len(name)):
                if "/" in name[-(i)]:
                    name = "File: " + name[-(i-1):]
                    break

            # Update Folder Name Label
            file_label.config(text = name)

    # Main Function
    def main(file_name):
        update_progress(1, "Harvesting files...")

        # Load Excel File
        old_path = str(file_name)
        old_book = load_workbook(old_path)  #read file
        old_sheet = old_book.active         #read sheet

        # Create a workbook and worksheet
        new_book = Workbook()               #create workbook
        new_sheet = new_book.active         #create worksheet

        # Column Headers
        headers =   ["open_access", "url", "title", "title_alternative", "doi", "source_fulltext_url",
                     "additional_text_uri", "author_classification", "total_author_count", "faculty_author_count", 
                     "institution_name", "authorized_name", "abstract", "meeting_name", "department1", "department2", 
                     "department3", "department4", "centers_labs", "centers_labs2", "centers_labs3", "centers_labs4", 
                     "pub_status", "comments", "keywords", "geographic_coverage", "time_period", "isbn", "issn", 
                     "document_type", "document_version", "file_type", "language_iso", "language2", "rights", 
                     "distribution_license", "publication_date", "custom_publication_date", "publisher",
                     "publisher_place", "source_publication", "volnum", "issnum", "articlenum", "fpage", 
                     "lpage", "pubmedid", "disciplines", "embargo_date", "date_uploaded", 
                     "primary_document_attached", "copyright", "author"]

        # Add Column Headers (before authors)
        new_col = 1
        for header in headers:
            new_sheet.cell(1,  new_col).value = header
            new_col += 1

        # Determine how many columns are in the sheet that is being read from
        old_max_col = len([c for c in old_sheet.iter_cols(min_row=1, max_row=1, values_only=True) if c[0] is not None])

        # Determine how many rows are in the sheet that is being read from
        old_max_row = len([c for c in old_sheet.iter_rows(min_col=1, max_col=1, values_only=True) if c[0] is not None])
        new_max_row = old_max_row

        # Starting author count
        last_cell = str(old_sheet.cell(1, old_max_col).value)[6:]

        index = 0
        for letter in last_cell:
            if letter == '_':
                author_count = int(last_cell[:index])
                break
            index += 1

        #Code for getting rid of empty author columns
        #temp_start = old_max_col + 1
        #is_used = False

        #while(not(is_used)):
            #update variables
        #    temp_stop = temp_start
        #    temp_start = temp_stop - 7

            #look for content
        #    for row in range(2, old_max_row + 1):
        #        for temp_col in range(temp_start, temp_stop): #searches through column headers
        #            if old_sheet.cell(row, temp_col).value:
        #                is_used = True
        #                print("Row: " + str(row) + " Col: " + str(temp_col) + " Cell value: " + str(old_sheet.cell(row, temp_col).value))
            
        #    print("Author " + str(author_count) + " " + str(is_used))
        #    if is_used == False:
        #        author_count -= 1

        print("Author Count: " + str(author_count))

        # Adds the 7 headers for an author
        def add_author_headers(new_col, num): #takes in starting column and author number
            #add headers to worksheet
            new_sheet.cell(1,  new_col).value = f'author{num}_fname'
            new_sheet.cell(1,  new_col+1).value = f'author{num}_mname'
            new_sheet.cell(1,  new_col+2).value = f'author{num}_lname'
            new_sheet.cell(1,  new_col+3).value = f'author{num}_suffix'
            new_sheet.cell(1,  new_col+4).value = f'author{num}_email'
            new_sheet.cell(1,  new_col+5).value = f'author{num}_institution'
            new_sheet.cell(1,  new_col+6).value = f'author{num}_is_corporate'

        # Add authors
        #print("Author Count: " + str(author_count))
        for num in range(1, author_count + 1):      #creates column headers for author count
            add_author_headers(new_col, num)        #add column headers for another author
            new_col += 7                            #update the current empty column header
        #    print("Author " + str(num) + " added.")

        # Determine how many columns are in the sheet that is being written to
        new_max_col = len([c for c in new_sheet.iter_cols(min_row=1, max_row=1, values_only=True) if c[0] is not None])

        # Creates yellow background fill
        yellow = PatternFill(patternType = 'solid', fgColor = 'fffb00')

        def new_index(new_header):
            for new_col in range(1, new_max_col + 1):
                new_col_head = new_sheet.cell(1, new_col).value
                if new_header == new_col_head:
                    return new_col

        def old_index(old_header):
            for old_col in range(1, old_max_col + 1):
                old_col_head = old_sheet.cell(1, old_col).value
                if old_header == old_col_head:
                    return old_col

        # Copy Information
        # col is the column index where info is being copied to (new excel)
        # old_col is the column index where info is being copied from (old excel)
        def copy(new_col, old_col):
            new_sheet.cell(row, new_col).value = old_sheet.cell(row, old_col).value

        # Fill Information
        # col is the column index where info is being filled (new excel)
        # fill_text is the text that is filled into the new excel
        def fill(new_col, fill_text):
            new_sheet.cell(row, new_col).value = fill_text

        # Add Highlight to Column Headers
        new_sheet.cell(1, new_index('open_access')).fill = yellow
        new_sheet.cell(1, new_index('url')).fill = yellow
        new_sheet.cell(1, new_index('copyright')).fill = yellow

        # * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
        #          FILLS ROW INFORMATION (in order of reformatted excel)            #
        # * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
        def row_information(row):
            # open_access
            try:
                if old_sheet.cell(row, old_index('Open Access')).value == "OA":
                    #Finds which specific OA it is
                    if old_sheet.cell(row, old_index('Gold OA')).value == "OA":
                        copy(new_index('open_access'), old_index('Gold OA'))
                    elif old_sheet.cell(row, old_index('Hybrid Gold OA')).value != "-":
                        copy(new_index('open_access'), old_index('Hybrid Gold OA'))
                    elif old_sheet.cell(row, old_index('Bronze OA')).value != "-":
                        copy(new_index('open_access'), old_index('Bronze OA'))
                    elif old_sheet.cell(row, old_index('Green Final OA')).value != "-":
                        copy(new_index('open_access'), old_index('Green Final OA'))
                    elif old_sheet.cell(row, old_index('Green Accepted OA')).value != "-":
                        copy(new_index('open_access'), old_index('Green Accepted OA')) 
                else:
                    #default OA
                    copy(new_index('open_access'), old_index('Open Access'))
            except TypeError:
                if row == 2:
                    warning_popup("Couldn't transfer 'open_access' column. The column may not exist.")

            # url
            try:
                copy(new_index('url'), old_index('source_fulltext_url'))
            except TypeError:
                if row == 2:
                    warning_popup("Couldn't transfer 'url' column. The column may not exist.")

            # title
            try:
                copy(new_index('title'), old_index('title'))
            except TypeError:
                if row == 2:
                    warning_popup("Couldn't transfer 'title' column. The column may not exist.")

            # doi
            try:
                copy(new_index('doi'), old_index('doi'))
            except TypeError:
                if row == 2:
                    warning_popup("Couldn't transfer 'doi' column. The column may not exist.")

            # source_fulltext_url
            try:
                copy(new_index('source_fulltext_url'), old_index('source_fulltext_url'))
            except TypeError:
                if row == 2:
                    warning_popup("Couldn't transfer 'source_fulltext_url' column. The column may not exist.")

            # author_classification
            fill(new_index('author_classification'), 'faculty')

            # total_author_count
            a_count = 0

            for c in range(old_index('author1_fname'), old_max_col + 1, 7): #searches through column headers
                if old_sheet.cell(row, c).value or old_sheet.cell(row, c + 1).value or old_sheet.cell(row, c + 2).value or old_sheet.cell(row, c + 3).value or old_sheet.cell(row, c + 4).value or old_sheet.cell(row, c + 5).value or old_sheet.cell(row, c + 6).value:
                    a_count += 1                                            #updates a_count
                else:
                    break

            fill(new_index('total_author_count'), a_count)

            # institution_name
            fill(new_index('institution_name'), 'Missouri University of Science and Technology')

            # abstract
            try:
                copy(new_index('abstract'), old_index('abstract'))
            except TypeError:
                if row == 2:
                    warning_popup("Couldn't transfer 'abstract' column. The column may not exist.")

            # Funding Number & Funding Sponsor
            try:
                f_num = old_sheet.cell(row, old_index('Funding Number')).value # Changing to grant?
                f_spon = old_sheet.cell(row, old_index('Funding Sponsor')).value # Changing to fundref?
                output = ''

                if f_spon != None:
                    if 'undefined' not in str(f_num):
                        output = str(f_spon) + ', Grant ' + str(f_num)
                    else:
                        output = str(f_spon)

                fill(new_index('comments'), output)
            except:
                if row == 2:
                    warning_popup("Couldn't transfer 'Funding Number' and 'Funding Sponsor' columns. The columns may not exist.")

            # keywords
            try:
                copy(new_index('keywords'), old_index('keywords'))
            except TypeError:
                if row == 2:
                    warning_popup("Couldn't transfer 'keywords' column. The column may not exist.")

            # isbn format function
            def format_isbn(val_in):
                if len(val_in) > 12 & len(val_in) < 16:         #between 13 and 15
                    if "[" in val_in:
                        string = val_in[2:15]                   #get number without [' ']
                        upper = string[:3]                      #get first 3 numbers
                        mid = string[3:12]                      #get middle 9 numbers
                        lower = string[12:]                     #get last number
                    else:
                        upper = val_in[:3]                      #get first 3 numbers
                        mid = val_in[3:12]                      #get middle 9 numbers
                        lower = val_in[12:]                     #get last number
                    output = upper + '-' + mid + '-' + lower    #format output
                    return output

            # isbn
            if old_sheet.cell(row, old_index('isbn')).value:
                val_in = str(old_sheet.cell(row, old_index('isbn')).value)
                if ',' in val_in:                               #two numbers
                    for char in val_in:
                        if ',' in char:
                            char = str(char)
                            first_num = val_in[:char]
                            second_num = val_in[char + 1:]
                            first_num = format_isbn(str(first_num))
                            second_num = format_isbn(str(second_num))
                            fill(new_index('isbn'), str(first_num) + ";" + str(second_num)) #output to cell
                            break
                else:                                           #one number
                   fill(new_index('isbn'), format_isbn(val_in)) #output to cell

            # issn
            e_ISSN = False
            issn = False

            if old_sheet.cell(row, old_index('e-ISSN')).value:
                val_in = old_sheet.cell(row, old_index('e-ISSN')).value
                string = str(val_in).zfill(8)                       #add missing leading 0s
                upper = string[:4]                                  #get first 4 numbers
                lower = string[4:]                                  #get last 4 numbers
                val_out = upper + "-" + lower                       #format output ####-####
                e_ISSN = True

            if old_sheet.cell(row, old_index('issn')).value:
                val_in2 = old_sheet.cell(row, old_index('issn')).value
                string2 = str(val_in2).zfill(8)                     #add missing leading 0s
                upper2 = string2[:4]                                #get first 4 numbers
                lower2 = string2[4:]                                #get last 4 numbers
                val_out2 = upper2 + "-" + lower2                    #format output ####-####
                issn = True

            output = ''

            if issn and e_ISSN:
                output = val_out2 + '; ' + val_out                  #outputs both issn
            else:
                if e_ISSN:
                    output = val_out                                #outputs 1st issn

                if issn:
                    output = val_out2                               #outputs 2nd issn

            fill(new_index('issn'), output)      #output to cell

            # document_type
            try:
                copy(new_index('document_type'), old_index('document_type'))
            except TypeError:
                if row == 2:
                    warning_popup("Couldn't transfer 'document_type' column. The column may not exist.")

            if 'conference' in str(new_sheet.cell(row, new_index('document_type')).value):
                fill(new_index('document_type'), 'article_conference_proceedings')

            # document_version
            fill(new_index('document_version'), 'Citation')

            # file_type
            fill(new_index('file_type'), 'text')

            # language_iso
            fill(new_index('language_iso'), 'English')

            # language2
            try:
                copy(new_index('language2'), old_index('language2'))
            except TypeError:
                if row == 2:
                    warning_popup("Couldn't transfer 'language2' column. The column may not exist.")

            # rights
            fill(new_index('rights'), '© 2021 , All rights reserved.')

            # publication_date
            try:
                copy(new_index('publication_date'), old_index('publication_date'))
            except TypeError:
                if row == 2:
                    warning_popup("Couldn't transfer 'publication_date' column. The column may not exist.")

            try:
                if new_sheet.cell(row, new_index('publication_date')).value:
                    date = str(new_sheet.cell(row, new_index('publication_date')).value)
                    if len(date) > 10:
                        date = date[:len(date) - 9] #gets rid of time in the format
                        fill(new_index('publication_date'), date)
            except:
                print("")

            # custom_publication_date
            try:
                if new_sheet.cell(row, new_index('publication_date')).value:
                    date = str(new_sheet.cell(row, new_index('publication_date')).value)
                    date = date[:10] #gets rid of time in the format

                    #WORKS WHEN FORMAT IS "2020-02-01 00:00:00"
                    # Get parts from date
                    year = date[:4]
                    month = date[5:7]
                    day = date[8:]

                    # Get month in letters
                    x = datetime.datetime(int(year), int(month), int(day))
                    month = x.strftime("%b")

                    # Format output
                    output = day + ' ' + month + ' ' + year

                    # Output
                    fill(new_index('custom_publication_date'), output)
            except:
                print ("Unusual formatting in 'publication_date' column")

            # source_publication
            filled = False

            # Open Known Meetings File
            path = 'R:/storage/libarchive/b/1. Processing/8. Other Projects/Scholars-Mine-GitHub/Control Panel/CODE/Harvesting/Text Files/KnownMeetings.txt'
            with open(path, "r") as all_meetings:
                # Check for matching meetings
                for one_meeting in all_meetings:
                    check = str(old_sheet.cell(row, old_index('source_publication')).value)
                    if check in str(one_meeting):
                        fill(new_index('meeting_name'), str(one_meeting))
                        if 'Proceedings of the ' not in check:
                            filled = True
                            fill(new_index('source_publication'), 'Proceedings of the ' + check)
                        break

            if filled == False:
                try:
                    copy(new_index('source_publication'), old_index('source_publication'))
                except TypeError:
                    if row == 2:
                        warning_popup("Couldn't transfer 'source_publication' column. The column may not exist.")

            # volnum
            try:
                copy(new_index('volnum'), old_index('volnum'))
            except TypeError:
                if row == 2:
                    warning_popup("Couldn't transfer 'volnum' column. The column may not exist.")

            # issnum
            try:
                copy(new_index('issnum'), old_index('issnum'))
            except TypeError:
                if row == 2:
                    warning_popup("Couldn't transfer 'issnum' column. The column may not exist.")

            # articlenum
            try:
                copy(new_index('articlenum'), old_index('articlenum'))
            except TypeError:
                if row == 2:
                    warning_popup("Couldn't transfer 'articlenum' column. The column may not exist.")

            # fpage
            try:
                copy(new_index('fpage'), old_index('fpage'))
            except TypeError:
                if row == 2:
                    warning_popup("Couldn't transfer 'fpage' column. The column may not exist.")

            # lpage
            try:
                copy(new_index('lpage'), old_index('lpage'))
            except TypeError:
                if row == 2:
                    warning_popup("Couldn't transfer 'lpage' column. The column may not exist.")

            # pubmedid
            try:
                copy(new_index('pubmedid'), old_index('pubmedid'))
            except TypeError:
                if row == 2:
                    warning_popup("Couldn't transfer 'pubmedid' column. The column may not exist.")
                try:
                    copy(new_index('pubmedid'), old_index('PubMed ID'))
                except TypeError:
                    if row == 2:
                        warning_popup("Couldn't transfer 'PubMed ID' column either. The column may not exist.")

            # primary_document_attached
            fill(new_index('primary_document_attached'), 'no')

            # copy author information
            for num in range(1, a_count + 1):               #copy author information
                copy(new_index(f'author{num}_fname'), old_index(f'author{num}_fname'))
                copy(new_index(f'author{num}_mname'), old_index(f'author{num}_mname'))
                copy(new_index(f'author{num}_lname'), old_index(f'author{num}_lname'))
                copy(new_index(f'author{num}_suffix'), old_index(f'author{num}_suffix'))
                copy(new_index(f'author{num}_email'), old_index(f'author{num}_email'))
                copy(new_index(f'author{num}_institution'), old_index(f'author{num}_institution'))
                copy(new_index(f'author{num}_is_corporate'), old_index(f'author{num}_is_corporate'))

        # Transfer row information for the whole sheet
        update_progress(2, "Harvesting files...") # Filling in " + str(old_max_row) + " rows

        for row in range(2, old_max_row + 1):
            row_information(row)

        # Fix All Uppercase Titles
        update_progress(3, "Harvesting files...") # Fixing all uppercase titles

        for row in range(2, new_max_row):
            title = new_sheet.cell(row, new_index('title')).value
            fill(new_index('title'), manual_upper(title))

        # Fix Subscripts and
        update_progress(4, "Harvesting files...") # Fixing subscripts and superscripts

        for row in range(2, new_max_row):
            fill(new_index('title'), replace_sub_sup(new_sheet.cell(row, new_index('title')).value))

        # Switching All Dates to Numbers
        update_progress(5, "Harvesting files...") # Switching all dates to numbers

        for row in range(2, new_max_row):
            # Fix issnum column
            col_value = str(new_sheet.cell(row, new_index('issnum')).value)
            if col_value:                                                   #if a value exists
                if '00:00:00' in col_value:                                 #change date to number
                    fill(new_index('issnum'), date_to_num(str(col_value)))
                elif '-' in col_value:                                      #get rid of dash
                    fill(new_index('issnum'), remove_dash(str(col_value)))

                if 'a' in col_value or 'e' in col_value or 'i' in col_value or 'o' in col_value or 'u' in col_value or 'y' in col_value: #get rid of words
                    fill(new_index('issnum'), '')

            # Fix volnum column
            col_value = str(new_sheet.cell(row, new_index('volnum')).value)
            if col_value:                                                   #if a value exists
                if '00:00:00' in col_value:                                 #change date to number
                    fill(new_index('volnum'), date_to_num(str(col_value)))
                elif '-' in col_value:                                      #get rid of dash
                    fill(new_index('volnum'), remove_dash(str(col_value)))

                if 'a' in col_value or 'e' in col_value or 'i' in col_value or 'o' in col_value or 'u' in col_value or 'y' in col_value: #get rid of words
                    fill(new_index('volnum'), '')

    #    # Database Check
    #    update_progress(6, "Harvesting files...") # Referencing database
    #    print()
    #    print('\t...referencing database (5/7)...')
    #
    #    for row in range(2, new_max_row + 1):
    #        total_author_count = new_sheet.cell(row, new_index('total_author_count')).value
    #        print('The total author count for row ' + str(row) + ' is ' + str(total_author_count))
    #        print()
    #
    #        for num in range (1, total_author_count + 1):
    #            author = []
    #            author.append(str(new_sheet.cell(row, new_index(f'author{num}_fname')).value))
    #            author.append(str(new_sheet.cell(row, new_index(f'author{num}_mname')).value))
    #            author.append(str(new_sheet.cell(row, new_index(f'author{num}_lname')).value))
    #            author.append(str(new_sheet.cell(row, new_index(f'author{num}_suffix')).value))
    #            author.append(str(new_sheet.cell(row, new_index(f'author{num}_email')).value))
    #            author.append(str(new_sheet.cell(row, new_index(f'author{num}_institution')).value))
    #            author.append(str(new_sheet.cell(row, new_index(f'author{num}_is_corporate')).value))
    #
    #            #print(author)
    #            #print()
    #
    #
    #            print('Email before: ' + str(author[4]))
    #            print()
    #
    #            # Identify S&T Faculty
    #            if(author[5] == 'Missouri University of Science and Technology' and author[4] is not None):
    #                author[4] = search_database(author[2], author[0])
    #
    #            if author[4] and author[5] == 'Missouri University of Science and Technology':
    #                print('Email after: ' + str(author[4]))
    #                print()
    #                fill(new_index('author' + str(num) + '_email'), author[4])
    #                print('Email added to row: ' + str(row) + '\t\tAuthor #: ' + str(num))
    #                print()
    #
    #            # Remove institution for any that aren't S&T (this includes co-authors and s&t students)

        # Prepare for Author_Split
        update_progress(7, "Harvesting files...") # Saving

        def author_split_information(row):
            upper = new_sheet.cell(row, new_index("total_author_count")).value

            for num in range(1, upper + 1):

                # gather name information
                last_name = new_sheet.cell(row, new_index("author" + str(num) + "_lname")).value
                first_name = new_sheet.cell(row, new_index("author" + str(num) + "_fname")).value
                

                if new_sheet.cell(row, new_index("author" + str(num) + "_mname")).value: #w/ middle name
                    middle_name = new_sheet.cell(row, new_index("author" + str(num) + "_mname")).value
                    middle_initial = middle_name[0:1]

                    if num > 1:
                        author_split = str(new_sheet.cell(row, new_index("author")).value) + " and " + str(last_name) + ", " + str(first_name) + " " + str(middle_initial) + "."
                    else:
                        author_split = str(last_name) + ", " + str(first_name) + " " + str(middle_initial) + "."
                else: #w/o middle name
                    if num > 1:
                        author_split = str(new_sheet.cell(row, new_index("author")).value) + " and " + str(last_name) + ", " + str(first_name)
                    else:
                        author_split = str(last_name) + ", " + str(first_name)

                try:
                    # author_split
                    fill(new_index("author"), str(author_split))
                except TypeError:
                    if row == 2:
                        warning_popup("Couldn't transfer 'author' column. The column may not exist.")

        for row in range(2, new_max_row + 1):
            author_split_information(row)

        # Format Columns
        update_progress(8, "Harvesting files...") # Adjusting column width

        for new_col in range(1, new_max_col + 1):
            max_length = 0
            for new_row in range(1, new_max_row):                               #go through columns and rows
                if len(str(new_sheet.cell(new_row, new_col).value)) > max_length:
                    max_length = len(str(new_sheet.cell(new_row, new_col).value))   #find the longest cell
                    if max_length >= 100:                                       #cell cannont be longer than 100
                        new_sheet.column_dimensions[get_column_letter(new_col)].width = 100
                        break

            # Adjust the cell length
            if max_length < 100:
                new_sheet.column_dimensions[get_column_letter(new_col)].width = max_length * 1.13

        # Save excel
        update_progress(9, "Harvesting files...") # Saving

        new_path = str(file_name)[:len(file_name) - 5] + '_ReadyForAuthorSplit.xlsx'
        new_book.save(new_path)

        update_progress(10, "Excel Created")

    # Start Button Function
    def start():
        # Run program and update progress bar
        try:
            main(harvesting.filename)
        except PermissionError:
            error_popup("Could not save processed excel. The excel file that needs to be saved over is currently open. Close the excel file and hit the 'start' button again.")
        #except:
        #    error_popup("There was an unknown error, the file could not be processed.")

    # Create a button
    browse_button = Button(frame, text = "Browse", command = lambda : browse(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    start_button = Button(frame, text = "Start", command = lambda : start(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")

    # Place in frame
    browse_button.grid(row = 0, column = 0, padx = (15, 0), pady = 15)
    start_button.grid(row = 0, column = 1, padx = 15, pady = 15)
