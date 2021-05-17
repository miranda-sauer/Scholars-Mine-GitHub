# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
#           IMPORT                                                                  #
# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
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
    open_all.configure()
    open_all.iconbitmap('R:/storage/libarchive/b/1. Processing/8. Other Projects/Control Panel/CODE/S&T_Logo.ico')

    # Center the window on the screen
    window_w = 396 # window width
    window_h = 310 # window height

    screen_w = open_all.winfo_screenwidth() # screen width
    screen_h = open_all.winfo_screenheight() # screen height

    x_coor = (screen_w/2) - (window_w/2) #middle of screen x coordinate
    y_coor = (screen_h/2) #lower half of screen y coordinate

    open_all.geometry("%dx%d+%d+%d" % (window_w, window_h, x_coor, y_coor)) # place in middle

    # Create Notebook and pack onto window
    notebook = ttk.Notebook(open_all)
    notebook.pack(padx = (2, 0))

    # Create Frames and pack onto window
    program = Frame(notebook, bg = '#003B49')
    error = Frame(notebook, bg = '#003B49')
    program.pack(fill = "both", expand = "1")
    error.pack(fill = "both", expand = "1")

    # Add frame to notebook
    notebook.add(program, text = "Program")
    notebook.add(error, text = "Error Log")

    # Create Progress Bar for program tab
    progress = ttk.Progressbar(program, orient = HORIZONTAL, length = 370, mode = 'determinate')

    # Create a control frame for program tab
    control = LabelFrame(program, text = "Select an Input File", bg = '#003B49', fg = 'white', font = 'tungsten 12 bold')

    # Create a label for program tab
    task = Label(program, text = "Waiting for a file", bg = '#003B49', fg = 'white', font = 'tungsten 12 bold')
    file_label = Label(program, text = "Folder: N/A", bg = '#003B49', fg = 'white', font = 'tungsten 12 bold')

    # Open Help Function
    def open_help():
        # Open word document
        os.startfile("R:/storage/libarchive/b/1. Processing/8. Other Projects/Control Panel/Documentation/Open All Help.docx")

    # Create a button for program tab
    help_button = Button(program, text = "Help", command = lambda : open_help(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    exit_button = Button(program, text = "Exit Open All", command = open_all.destroy, bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")

    # Place in window for program tab
    control.pack(padx = 20, pady = 10)
    file_label.pack(padx = 20, pady = (0, 20))
    progress.pack(padx = 5, pady = 5)
    task.pack(padx = 20)
    help_button.pack(padx = (331, 10), pady = 0)
    exit_button.pack(padx = (270, 10), pady = (5, 10))

    # Create Error Log
    def update_error_log(file_name, error_message):
        current_time = datetime.datetime.now()
        path = "R:/storage/libarchive/b/1. Processing/8. Other Projects/Control Panel/CODE Version 1.2/Harvesting/Text Files/ErrorLog.txt"
        with open(path, "a+") as el: #Open error log to append
            el.write("\n" + str(current_time) + "\tOpen All : " + str(file_name) + " : " + str(error_message)) #Append Error Log

    # Create & place a recent error frame for error tab
    recent_error_frame = LabelFrame(error, text = "Recent Error Log", bg = '#003B49', fg = 'white', font = 'tungsten 12 bold')
    recent_error_frame.pack(padx = 10, pady = 10)

    # Create & place a Label for recent error frame
    recent_error_label = Label(recent_error_frame, text = "Folder: N/A\tErrors: N/A", bg = '#003B49', fg = 'white', font = 'tungsten 12 bold')
    recent_error_label.pack(padx = 10, pady = 10)

    # Open Error Log
    def open_error_log():
        # Open word document
        os.startfile("R:/storage/libarchive/b/1. Processing/8. Other Projects/Control Panel/CODE Version 1.2/Harvesting/Text Files/ErrorLog.txt")

    # Create & place a button for error tab
    error_log_button = Button(error, text = "Open Error Log", command = lambda : open_error_log(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    error_log_button.pack(pady = 10)

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
        update_progress(0, "Waiting for a file")

        # Open file
        open_all.filename = filedialog.askdirectory(initialdir = "R:/storage/libarchive/a/Student Processing", title = "Select Input",)

        # Get file name
        name = open_all.filename
        for i in range(len(name)):
            if "/" in name[-(i)]:
                name = "Folder: " + name[-(i-1):]
                break

        # Update Folder Name Label
        file_label.config(text = name)

        # Update Error Report Label
        recent_error_label.config(text = name + "\tError: N/A")

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

        try:
            fail = 3/0
        except:
            print("purposeful error")
            update_error_log(name, "purposeful error")

    # Start Button Function
    def start():
        # Run program and update progress bar
        main(open_all.filename)
    
    # Create a button for program tab
    browse_button = Button(control, text = "Browse", command = lambda : browse(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    start_button = Button(control, text = "Start", command = lambda : start(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    
    # Place in control frame for program tab
    browse_button.grid(row = 0, column = 0, padx = (15, 0), pady = 15)
    start_button.grid(row = 0, column = 1, padx = 15, pady = 15)