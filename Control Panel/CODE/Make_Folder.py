# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
#           IMPORT                                                                  #
# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
import time
import os
import shutil



# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
#           DISPLAY WINDOW                                                          #
# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #

def open_make_folder():
    # Create window
    make_folder = Toplevel()
    make_folder.title("Make Folder Program")
    make_folder.configure(bg = '#003B49')
    make_folder.iconbitmap('R:/storage/libarchive/b/1. Processing/8. Other Projects/Scholars-Mine-GitHub/Control Panel/CODE/S&T_Logo.ico')

    # Center the window on the screen
    window_w = 400 # window width
    window_h = 290 # window height

    screen_w = make_folder.winfo_screenwidth() # screen width
    screen_h = make_folder.winfo_screenheight() # screen height

    x_coor = (screen_w/2) - (window_w/2) #middle of screen x coordinate
    y_coor = (screen_h/2) #lower half of screen y coordinate

    make_folder.geometry("%dx%d+%d+%d" % (window_w, window_h, x_coor, y_coor)) # place in middle

    #Create Progress Bar
    progress = ttk.Progressbar(make_folder, orient = HORIZONTAL, length = 370, mode = 'determinate')

    # Create a frame
    frame = LabelFrame(make_folder, text = "Select an Input File", bg = '#003B49', fg = 'white', font = 'tungsten 12 bold')

    # Create a label
    task = Label(make_folder, text = "Waiting for a file", bg = '#003B49', fg = 'white', font = 'tungsten 12 bold')
    file_label = Label(make_folder, text = "Folder: N/A", bg = '#003B49', fg = 'white', font = 'tungsten 12 bold')


    # Open Help Function
    def open_help():
        # Open word document
        os.startfile("R:/storage/libarchive/b/1. Processing/8. Other Projects/Control Panel/Documentation/Make Folder Help.docx")

    # Create a button
    help_button = Button(make_folder, text = "Help", command = lambda : open_help(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    exit_button = Button(make_folder, text = "Exit Make Folder", command = make_folder.destroy, bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")

    # Place in window
    frame.pack(padx = 20, pady = 10)
    file_label.pack(padx = 20, pady = (0, 20))
    progress.pack(padx = 5, pady = 5)
    task.pack(padx = 20)
    help_button.pack(padx = (331, 10), pady = 0)
    exit_button.pack(padx = (240, 10), pady = (5, 0))

    # Update Progress Bar
    def update_progress(p, t):  
        # Update bar value
        progress['value'] = (p/2)*100

        # Update bar label
        task.config(text = t)
        task.pack() 

        # Refresh window (very important line!)
        make_folder.update_idletasks()
        time.sleep(0.2)

    # Function Click
    def browse():
        # Reset progress bar
        update_progress(0, "Waiting for a file")

        # Open file
        make_folder.filename = filedialog.askdirectory(initialdir = "R:/storage/libarchive/a/Student Processing", title = "Select Input",)

        #Get file name
        name = make_folder.filename
        for i in range(len(name)):
            if "/" in name[-(i)]:
                name = "Folder: " + name[-(i-1):]
                break

        #Update Folder Name Label
        file_label.config(text = name)

    # Main Function
    def main(file_name):
        update_progress(1, "Making Folder...")

        path = file_name
        os.chdir(path)

        for file in os.listdir() :
            name, ext = os.path.splitext(file)
            target = f'{path}/{name}'
            try:
                os.mkdir(target)
            except: 
                pass
            shutil.move(file, target)

        update_progress(2, "Program Complete")
    
    # Start Button Function
    def start():
        # Run program and update progress bar
        main(make_folder.filename)
    
    # Create a button
    browse_button = Button(frame, text = "Browse", command = lambda : browse(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    start_button = Button(frame, text = "Start", command = lambda : start(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")

    # Place in frame
    browse_button.grid(row = 0, column = 0, padx = (15, 0), pady = 15)
    start_button.grid(row = 0, column = 1, padx = 15, pady = 15)