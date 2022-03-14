# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
#           IMPORT                                                                  #
# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
import datetime
import time
import os
import glob
import subprocess



# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
#           DISPLAY WINDOW                                                          #
# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #

def open_open_all():
    # Create window
    open_all = Toplevel()
    open_all.title("Open All Program")
    open_all.configure(bg = '#003B49')
    open_all.iconbitmap('R:/storage/libarchive/a/Scholars-Mine-GitHub/Control Panel/CODE/S&T_Logo.ico')

    # Center the window on the screen
    window_w = 400 # window width
    window_h = 290 # window height

    screen_w = open_all.winfo_screenwidth() # screen width
    screen_h = open_all.winfo_screenheight() # screen height

    x_coor = (screen_w/2) - (window_w/2) #middle of screen x coordinate
    y_coor = (screen_h/2) #lower half of screen y coordinate

    open_all.geometry("%dx%d+%d+%d" % (window_w, window_h, x_coor, y_coor)) # place in middle

    # Create Progress Bar
    progress = ttk.Progressbar(open_all, orient = HORIZONTAL, length = 370, mode = 'determinate')

    # Create a frame
    frame = LabelFrame(open_all, text = "Select an Input File", bg = '#003B49', fg = 'white', font = 'tungsten 12 bold')

    # Create a label
    task = Label(open_all, text = "Waiting for a file", bg = '#003B49', fg = 'white', font = 'tungsten 12 bold')
    file_label = Label(open_all, text = "Folder: N/A", bg = '#003B49', fg = 'white', font = 'tungsten 12 bold')

    # Open Help Function
    def open_help():
        # Open word document
        try:
            os.startfile("R:/storage/libarchive/a/Scholars-Mine-GitHub/Control Panel/Documentation/Open All Help.docx")
        except:
            error_popup("Could not open help file.")

    # Create a button
    help_button = Button(open_all, text = "Help", command = lambda : open_help(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    exit_button = Button(open_all, text = "Exit Open All", command = open_all.destroy, bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")

    # Place in window
    frame.pack(padx = 20, pady = 10)
    file_label.pack(padx = 20, pady = (0, 20))
    progress.pack(padx = 5, pady = 5)
    task.pack(padx = 20)
    help_button.pack(padx = (331, 10), pady = 0)
    exit_button.pack(padx = (270, 10), pady = (5, 10))

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
        open_all.update_idletasks()
        time.sleep(0.2)

    # Function Click
    def browse():
        # Reset progress bar
        update_progress(0, "Waiting for a folder")

        # Open file
        open_all.filename = filedialog.askdirectory(initialdir = "R:/storage/libarchive/a/Student Processing", title = "Select Input",)
        if open_all.filename == "":
            warning_popup("No folder selected.")
            file_label.config(text = "Folder: N/A")
            del open_all.filename
        else:
            # Get file name
            name = open_all.filename
            for i in range(len(name)):
                if "/" in name[-(i)]:
                    name = "Folder: " + name[-(i-1):]
                    break

            # Update Folder Name Label
            file_label.config(text = name)

    # Main Function
    def main(file_name):
        update_progress(1, "Opening all...")

        path = file_name
        os.chdir(path)
        fileList = glob.glob('**/*', recursive=True)
  
        for file in fileList:
            name, ext = os.path.splitext(file)
            if ext == '.pdf':
                subprocess.Popen([file], shell=True)  

        update_progress(2, "Program Complete")

    # Start Button Function
    def start():
        # Run program and update progress bar
        try:
            main(open_all.filename)
        except AttributeError:
            error_popup("No folder selected. Browse to select a folder.")
        except:
            error_popup("There was an unknown error, the folder could not be opened.")
    
    # Create a button for program tab
    browse_button = Button(frame, text = "Browse", command = lambda : browse(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    start_button = Button(frame, text = "Start", command = lambda : start(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    
    # Place in control frame for program tab
    browse_button.grid(row = 0, column = 0, padx = (15, 0), pady = 15)
    start_button.grid(row = 0, column = 1, padx = 15, pady = 15)