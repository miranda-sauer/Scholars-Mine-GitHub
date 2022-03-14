# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
#                                       Import                                      #
# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
import time
import os
import glob
import shutil

def open_zip_files():
    # * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
    #                      Variable Declaration and Initialization                      #
    # * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
    Zip = Toplevel() # Create window

    frame = Frame(Zip, bg = '#003B49') # Create a frame
    progress = ttk.Progressbar(Zip, orient = HORIZONTAL, length = 400, mode = 'determinate') # Create Progress Bar
    task = Label(Zip, text = "Waiting for a file", bg = '#003B49', fg = 'white', font = 'tungsten 12 bold') # Create a Label
    file_label = Label(Zip, text = "Folder: N/A", bg = '#003B49', fg = 'white', font = 'tungsten 12 bold') # Create a Label
    help_button = Button(Zip, text = "Help", command = lambda : open_help(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge") # Create a Button
    select_button = Button(Zip, text = "Select File(s)", command = lambda : browse(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge") # Create a Button




    # * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
    #                      Function Declaration and Initialization                      #
    # * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
    
    # Error Message Popup
    def error_popup(error_message):
        messagebox.showerror("Error", error_message)

    # Open Help Function
    def open_help():
        # Open word document
        try:
            os.startfile("R:/storage/libarchive/a/zzz_Programs/Control Panel/Documentation/Zip Files Help.docx")
        except:
            error_popup("Could not open help file.")



    # * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
    #                                      Main                                         #
    # * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
    # Customize window
    Zip.title('Zip Files Program')
    Zip.configure(bg = '#003B49')
    Zip.iconbitmap('R:/storage/libarchive/a/zzz_Programs/Control Panel/CODE/S&T_Logo.ico')

    # Place in window
    frame.pack(padx = 20, pady = 10)
    file_label.pack(padx = 20, pady = (0, 20))
    progress.pack(padx = 5, pady = 5)
    task.pack(padx = 20)
    help_button.pack(padx = (331, 10), pady = 0)
    exit_button.pack(padx = (205, 10), pady = (5, 0))

    # Place in frame
    help_button.pack(side = RIGHT, padx = 5)
    select_button.pack(side = LEFT, padx = 5)

    # Place in middle_frame
    scroll_bar.pack(side = RIGHT, fill = Y)
    label_canvas.pack(side = LEFT, fill = BOTH, expand = 1)

    # Place in bottom_frame
    progress.grid(row = 1, column = 0, padx = 5, pady = 5)
    start_button.grid(row = 1, column = 1, padx = 5, pady = 5)

    # Configure
    label_canvas.configure(yscrollcommand = scroll_bar.set)
    label_canvas.bind('<Configure>', lambda e: label_canvas.configure(scrollregion = label_canvas.bbox("all")))
    label_canvas.bind_all('<MouseWheel>', lambda e: label_canvas.yview('scroll', (-1*e.delta), "units"))

    # Place in Canvas
    label_canvas.create_window((0, 0), window = canvas_frame, anchor = 'nw')

    # Place in canvas_frame
    file_label.pack(fill = BOTH)

    # Keeps window open until closed
    Both.mainloop()