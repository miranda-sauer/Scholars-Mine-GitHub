# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
#           IMPORT                                                                  #
# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
import time
import os
import glob



# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #
#           DISPLAY WINDOW                                                          #
# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * #

def open_cannon_remake():
    # Create window
    cannon_remake = Toplevel()
    cannon_remake.title("Cannon Remake Program")
    cannon_remake.configure(bg = '#003B49')
    cannon_remake.iconbitmap('R:/storage/libarchive/b/1. Processing/8. Other Projects/Scholars-Mine-GitHub/Control Panel/CODE/S&T_Logo.ico')

    # Center the window on the screen
    window_w = 400 # window width
    window_h = 290 # window height

    screen_w = cannon_remake.winfo_screenwidth() # screen width
    screen_h = cannon_remake.winfo_screenheight() # screen height

    x_coor = (screen_w/2) - (window_w/2) #middle of screen x coordinate
    y_coor = (screen_h/2) #lower half of screen y coordinate
    
    cannon_remake.geometry("%dx%d+%d+%d" % (window_w, window_h, x_coor, y_coor)) # place in middle

    #Create Progress Bar
    progress = ttk.Progressbar(cannon_remake, orient = HORIZONTAL, length = 370, mode = 'determinate')

    # Create a frame
    frame = LabelFrame(cannon_remake, text = "Select an Input File", bg = '#003B49', fg = 'white', font = 'tungsten 12 bold')

    # Create a label
    task = Label(cannon_remake, text = "Waiting for a file", bg = '#003B49', fg = 'white', font = 'tungsten 12 bold')
    file_label = Label(cannon_remake, text = "Folder: N/A", bg = '#003B49', fg = 'white', font = 'tungsten 12 bold')

    # Open Help Function
    def open_help():
        # Open word document
        try:
            os.startfile("R:/storage/libarchive/b/1. Processing/8. Other Projects/Control Panel/Documentation/Cannon Remake Help.docx")
        except:
            error_popup("Could not open help file.")

    # Create a button
    help_button = Button(cannon_remake, text = "Help", command = lambda : open_help(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    exit_button = Button(cannon_remake, text = "Exit Cannon Remake", command = cannon_remake.destroy, bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")

    # Place in window
    frame.pack(padx = 20, pady = 10)
    file_label.pack(padx = 20, pady = (0, 20))
    progress.pack(padx = 5, pady = 5)
    task.pack(padx = 20)
    help_button.pack(padx = (331, 10), pady = 0)
    exit_button.pack(padx = (205, 10), pady = (5, 0))

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
        cannon_remake.update_idletasks()
        time.sleep(0.2)

    # Function Click
    def browse():
        # Reset progress bar
        update_progress(0, "Waiting for a folder")

        # Open file
        cannon_remake.filename = filedialog.askdirectory(initialdir = "R:/storage/libarchive/a/Student Processing", title = "Select Input",)
        if cannon_remake.filename == "":
            warning_popup("No folder selected.")
            file_label.config(text = "Folder: N/A")
            del cannon_remake.filename
        else:
            #Get file name
            name = cannon_remake.filename
            for i in range(len(name)):
                if "/" in name[-(i)]:
                    name = "Folder: " + name[-(i-1):]
                    break

            #Update Folder Name Label
            file_label.config(text = name)

    # Main Function
    def main(file_name):
        update_progress(1, "Running Cannon Remake...")

        path = cannon_remake.filename
        os.chdir(path)
        fileList = os.listdir()

        for file in fileList:
            name, ext = os.path.splitext(file)
            if '_Page_' in name:
                name_ = name.split('_Page_')
                newName = name_[0] + name_[1]
                # while len(name_[1]) < 4:
                #     name_[1] = '0' + name_[1]
                os.rename(file, f'{newName}{ext}')

        update_progress(2, "Program Complete")
    
    # Start Button Function
    def start():
        # Run program and update progress bar
        try:
            main(cannon_remake.filename)
        except AttributeError:
            error_popup("No folder selected. Browse to select a folder.")
        except:
            error_popup("There was an unknown error, the folder could not be remade.")
    
    # Create a button
    browse_button = Button(frame, text = "Browse", command = lambda : browse(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    start_button = Button(frame, text = "Start", command = lambda : start(), bg = '#78BE20', fg = '#003B49', font = 'tungsten 12 bold', borderwidth = 1, relief = "ridge")
    
    # Place in frame
    browse_button.grid(row = 0, column = 0, padx = (15, 0), pady = 15)
    start_button.grid(row = 0, column = 1, padx = 15, pady = 15)