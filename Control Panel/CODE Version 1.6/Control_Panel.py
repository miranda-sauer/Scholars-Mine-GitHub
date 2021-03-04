from tkinter import *
from tkinter import filedialog
from tkinter import ttk
import time
import os
from Harvesting import open_harvesting
from Cannon_Remake import open_cannon_remake
from Make_Folder import open_make_folder
from Open_All import open_open_all

def open_control_panel():

    # Create window
    root = Tk()
    root.title("Scholars' Mine Control Panel")
    root.configure(bg = '#003B49')
    root.iconbitmap('R:/storage/libarchive/b/1. Processing/8. Other Projects/Control Panel/CODE/S&T_Logo.ico')

    # Center the window on the screen
    window_w = 462 # window width
    window_h = 220 # window height

    screen_w = root.winfo_screenwidth() # screen width
    screen_h = root.winfo_screenheight() # screen height

    x_coor = (screen_w/2) - (window_w/2) #middle of screen x coordinate
    y_coor = (screen_h/2) - (window_h/2) #middle of screen y coordinate

    root.geometry("%dx%d+%d+%d" % (window_w, window_h, x_coor, y_coor)) # place in middle

    # Open Help Function
    def open_help():
        # Open word document
        os.startfile("R:/storage/libarchive/b/1. Processing/8. Other Projects/Control Panel/Documentation/Control Panel Help.docx")

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

    # Create a button
    author_split = Button(frame, text = "Author Split", bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    harvest = Button(frame, text = "Harvest", command = lambda : open_harvesting(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    cannon_remake = Button(frame, text = "Cannon Remake", command = lambda : open_cannon_remake(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    make_folder = Button(frame, text = "Make Folder", command = lambda : open_make_folder(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    open_all = Button(frame, text = "Open All", command = lambda : open_open_all(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    help_button = Button(root, text = "Help", command = lambda : open_help(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    exit_button = Button(root, text = "Exit All", command = root.quit, bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")

    # Place in frame
    author_split.grid(row = 0, column = 0, padx = 20, pady = 20)
    harvest.grid(row = 0, column = 1, padx = 0, pady = 20)
    cannon_remake.grid(row = 0, column = 2, padx = 20, pady = 20)
    make_folder.grid(row = 1, column = 0, padx = 20, pady = (0, 20))
    open_all.grid(row = 1, column = 1, padx = 0, pady = (0, 20))

    # Place in window
    frame.grid(row = 0, column = 0, columnspan = 2, padx = 24, pady = 10)
    help_button.grid(row = 1, column = 0, pady = (5, 0))
    exit_button.grid(row = 1, column = 1, pady = (5, 0))

    # Keeps window open until closed
    root.mainloop()