from tkinter import *
from tkinter import filedialog
from tkinter import ttk
import time
import os
from Harvesting import open_harvesting
from Cannon_Remake import open_cannon_remake
from Make_Folder import open_make_folder
from Open_All import open_open_all
from Diacritics import open_diacritics
from Reformat_Harvesting import open_reformat_harvesting

def open_control_panel():

    # Create window
    root = Tk()
    root.title("Scholars' Mine Control Panel")
    root.configure(bg = '#003B49')
    root.iconbitmap('R:/storage/libarchive/b/1. Processing/8. Other Projects/Scholars-Mine-GitHub/Control Panel/CODE/S&T_Logo.ico')

    # Center the window on the screen
    window_w = 618 # window width
    window_h = 265 # window height

    screen_w = root.winfo_screenwidth() # screen width
    screen_h = root.winfo_screenheight() # screen height

    x_coor = (screen_w/2) - (window_w/2) #middle of screen x coordinate
    y_coor = (screen_h/2) - ((3/2)*window_h) #upper half of screen y coordinate

    root.geometry("%dx%d+%d+%d" % (window_w, window_h, x_coor, y_coor)) # place in middle

    # Open Help Function
    def open_help():
        # Open word document
        os.startfile("R:/storage/libarchive/b/1. Processing/8. Other Projects/Control Panel/Documentation/Control Panel Help.docx")

    # Open Error Log Function
    def open_error_log():
        # Open .txt file
        os.startfile("R:/storage/libarchive/b/1. Processing/8. Other Projects/Scholars-Mine-GitHub/Control Panel/Documentation/Error Log.txt")

    # Create a frame
    frame = LabelFrame(root, text = "Select a Program", bg = '#003B49', fg = "white", font = 'tungsten 14 bold')

    # Open program functions
    def open_harvesting_program():
        open_harvesting()

    def open_cannon_remake_program():
        open_cannon_remake()

    def open_make_folder_program():
        open_make_folder()

    def open_open_all_program():
        open_open_all()

    def open_diacritics_program():
        open_diacritics()

    def open_reformat_harvesting_program():
        open_reformat_harvesting()


    # Create a button
    author_split = Button(frame, text = "      Author Split       ", bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    harvest = Button(frame, text = "            Harvest             ", command = lambda : open_harvesting(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    cannon_remake = Button(frame, text = "  Cannon Remake  ", command = lambda : open_cannon_remake(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    make_folder = Button(frame, text = "      Make Folder      ", command = lambda : open_make_folder(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    open_all = Button(frame, text = "            Open All            ", command = lambda : open_open_all(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    diacritics = Button(frame, text = "         Diacritics         ", command = lambda : open_diacritics(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    reformat_harvesting = Button(frame, text = "Reformat Harvesting ", command = lambda : open_reformat_harvesting(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    help_button = Button(root, text = "      Help      ", command = lambda : open_help(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    error_log_button = Button(root, text = "   Error Log   ", command = lambda : open_error_log(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    exit_button = Button(root, text = "     Exit All     ", command = root.quit, bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")

    # Place in frame
    author_split.grid(row = 0, column = 0, padx = (20,10), pady = (15,10))
    harvest.grid(row = 0, column = 1, padx = 10, pady = (15,10))
    cannon_remake.grid(row = 0, column = 2, padx = (10,20), pady = (15,10))
    make_folder.grid(row = 1, column = 0, padx = (20,10), pady = 5)
    open_all.grid(row = 1, column = 1, padx = 10, pady = 5)
    diacritics.grid(row = 1, column = 2, padx = (10,20), pady = 5)
    reformat_harvesting.grid(row = 2, column = 1, padx = 10, pady = (10,15))

    # Place in window
    frame.grid(row = 0, column = 0, columnspan = 3, padx = 24, pady = 15)
    help_button.grid(row = 1, column = 0, pady = 5)
    error_log_button.grid(row = 1, column = 1, pady = 5)
    exit_button.grid(row = 1, column = 2, pady = 5)

    # Keeps window open until closed
    root.mainloop()