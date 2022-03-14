from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
import datetime
import time
import os
import glob
import subprocess
#from UI_Author_Split import open_author_split
#from Harvesting import open_harvesting
from Cannon_Remake import open_cannon_remake
from Make_Folder import open_make_folder
from Open_All import open_open_all
from Both import open_both
from Zip_Files import open_zip_files

# Error Message Popup
def error_popup(error_message):
    messagebox.showerror("Error", error_message)

def open_control_panel():

    # Create window
    root = Tk()
    root.title("Scholars' Mine Control Panel")
    root.configure(bg = '#003B49')
    root.iconbitmap('R:/storage/libarchive/a/zzz_Programs/Control Panel/CODE/S&T_Logo.ico')

    # Center the window on the screen
    window_w = 618 # window width
    window_h = 220 # window height

    screen_w = root.winfo_screenwidth() # screen width
    screen_h = root.winfo_screenheight() # screen height

    x_coor = (screen_w/2) - (window_w/2) #middle of screen x coordinate
    y_coor = (screen_h/2) - ((3/2)*window_h) #upper half of screen y coordinate

    root.geometry("%dx%d+%d+%d" % (window_w, window_h, x_coor, y_coor)) # place in middle

    # Open Help Function
    def open_help():
        # Open word document
        try:
            os.startfile("R:/storage/libarchive/a/zzz_Programs/Control Panel/Documentation/Control Panel Help.docx")
        except:
            error_popup("Could not open help file.")

    # Create a frame
    frame = LabelFrame(root, text = "Select a Program", bg = '#003B49', fg = "white", font = 'tungsten 14 bold')

    # Open Programs
    #def open_author_split_program():
    #    try:
    #        open_author_split()
    #    except:
    #        error_popup("Could not open harvesting program.")

    #def open_harvesting_program():
    #    try:
    #        open_harvesting()
    #    except:
    #        error_popup("Could not open harvesting program.")

    def open_cannon_remake_program():
        try:
            open_cannon_remake()
        except:
            error_popup("Could not open cannon remake program.")

    def open_make_folder_program():
        try:
            open_make_folder()
        except:
            error_popup("Could not open make folder program.")

    def open_open_all_program():
        try:
            open_open_all()
        except:
            error_popup("Could not open open all program.")

    def open_both_program():
        try:
            open_both()
        except:
            error_popup("Could not open both program.")

    def open_zip_files_program():
        error_popup("The zip files program is under construction.")
        #try:
        #    open_zip_files()
        #except:
        #    error_popup("Could not open zip files program.")

    # Create a button
    #author_split = Button(frame, text = "        Author Split       ", command = lambda : open_author_split_program(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    #harvest = Button(frame, text = "Harvest & Reformat", command = lambda : open_harvesting_program(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    both = Button(frame, text = "             Both             ", command = lambda : open_both_program(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")

    make_folder = Button(frame, text = "       Make Folder       ", command = lambda : open_make_folder_program(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    open_all = Button(frame, text = "           Open All          ", command = lambda : open_open_all_program(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    cannon_remake = Button(frame, text = "  Cannon Remake  ", command = lambda : open_cannon_remake_program(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")

    zip_files = Button(frame, text = "           Zip Files          ", command = lambda : open_zip_files_program(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")

    help_button = Button(root, text = "      Help      ", command = lambda : open_help(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")

    # Place in frame
    both.grid(row = 0, column = 0, padx = (20,10), pady = (15,10))
    make_folder.grid(row = 0, column = 1, padx = 10, pady = (15,10))
    open_all.grid(row = 0, column = 2, padx = (10,20), pady = (15,10))
    
    cannon_remake.grid(row = 1, column = 0, padx = (20,10), pady = (5,20))
    zip_files.grid(row = 1, column = 1, padx = 10, pady = (5,20))

    #Both of these programs below have been replaced with the "Both" program, 
    # the code is still here if you need to run them seperately, 
    # just uncomment stuff, but the code may not be up to date with 
    # any changes that are in the both program
    
    #author_split.grid(row = 0, column = 0, padx = (20,10), pady = (15,10))
    #harvest.grid(row = 0, column = 2, padx = 10, pady = (15,10))

    # Place in window
    frame.grid(row = 0, column = 0, padx = 24, pady = 15)
    help_button.grid(row = 1, column = 0, pady = 5)

    # Keeps window open until closed
    root.mainloop()