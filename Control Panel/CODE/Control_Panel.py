from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
import datetime
import time
import os
import glob
import subprocess
from UI_Author_Split import open_author_split
from Harvesting import open_harvesting
from Cannon_Remake import open_cannon_remake
from Make_Folder import open_make_folder
from Open_All import open_open_all
from Diacritics import open_diacritics

# Error Message Popup
def error_popup(error_message):
    messagebox.showerror("Error", error_message)

def open_control_panel():

    # Create window
    root = Tk()
    root.title("Scholars' Mine Control Panel")
    root.configure(bg = '#003B49')
    root.iconbitmap('R:/storage/libarchive/b/1. Processing/8. Other Projects/Scholars-Mine-GitHub/Control Panel/CODE/S&T_Logo.ico')

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
            os.startfile("R:/storage/libarchive/b/1. Processing/8. Other Projects/Control Panel/Documentation/Control Panel Help.docx")
        except:
            error_popup("Could not open help file.")

    # Create a frame
    frame = LabelFrame(root, text = "Select a Program", bg = '#003B49', fg = "white", font = 'tungsten 14 bold')

    # Open Programs
    def open_author_split_program():
        try:
            open_author_split()
        except:
            error_popup("Could not open harvesting program.")

    def open_harvesting_program():
        try:
            open_harvesting()
        except:
            error_popup("Could not open harvesting program.")

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

    def open_diacritics_program():
        #try:
        #    open_diacritics()
        #except:
        #    error_popup("Could not open diacritics program.")
        error_popup("The diacritics program is not currently working.")

    # Create a button
    author_split = Button(frame, text = "        Author Split       ", command = lambda : open_author_split(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    harvest = Button(frame, text = "Harvest & Reformat", command = lambda : open_harvesting(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    diacritics = Button(frame, text = "         Diacritics         ", command = lambda : open_diacritics_program(), bg = '#DCE3E4', fg = '#63666A', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")

    make_folder = Button(frame, text = "       Make Folder       ", command = lambda : open_make_folder(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    open_all = Button(frame, text = "           Open All          ", command = lambda : open_open_all(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    cannon_remake = Button(frame, text = "  Cannon Remake  ", command = lambda : open_cannon_remake(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")

    help_button = Button(root, text = "      Help      ", command = lambda : open_help(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    exit_button = Button(root, text = "     Exit All     ", command = root.quit, bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")

    # Place in frame
    author_split.grid(row = 0, column = 0, padx = (20,10), pady = (15,10))
    harvest.grid(row = 0, column = 1, padx = 10, pady = (15,10))
    diacritics.grid(row = 0, column = 2, padx = (10,20), pady = (15,10))

    make_folder.grid(row = 1, column = 0, padx = (20,10), pady = (5,20))
    open_all.grid(row = 1, column = 1, padx = 10, pady = (5,20))
    cannon_remake.grid(row = 1, column = 2, padx = (10,20), pady = (5,20))

    # Place in window
    frame.grid(row = 0, column = 0, columnspan = 3, padx = 24, pady = 15)

    help_button.grid(row = 1, column = 0, pady = 5)
    exit_button.grid(row = 1, column = 2, pady = 5)

    # Keeps window open until closed
    root.mainloop()