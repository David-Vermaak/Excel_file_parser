import tkinter as tk
from tkinter import ttk
import tkinter.scrolledtext as st
import sv_ttk

#Regex import
import re
from tkinter import filedialog
import pandas as pd


def display_button():
    clear_window()
    open_button = ttk.Button(content_frame, text="Open Excel File", command=open_excel_file, style="Accent.TButton")
    open_button.pack(pady=20)


def open_excel_file():
    #this opens the choose file dialog
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        try:
            #converts the file to a data structure using pandas
            df = pd.read_excel(file_path)
            #fills the blank cells 'na' with empty strings ''
            df = df.fillna('')
            clear_window()
            display_data(df)

        except Exception as e:
            tk.messagebox.showerror("Error", f"An error occurred: {str(e)}")



def display_data(df):

    # Create a ttk.Treeview widget
    tree = ttk.Treeview(content_frame, columns=list(df.columns), show="headings")

    # Create horizontal scrollbar
    hsb = ttk.Scrollbar(content_frame, orient="horizontal", command=tree.xview)
    tree.configure(xscrollcommand=hsb.set)

    # Add column headings
    for column in df.columns:
        tree.heading(column, text=column)
        tree.column(column, width=100)  # Adjust the column width as needed

    # Insert data from the DataFrame into the Treeview
    for index, row in df.iterrows():
        tree.insert("", index, values=list(row))

    # Pack the Treeview and horizontal scrollbar
    tree.pack(fill="both", expand=True)
    hsb.pack(side="bottom", fill="x")







#information textbox
def info():

    clear_window()

    title = "Program Information"
    info_text = """
        Welcome to the Excel File Parser
        How to Use:

        1.) Select a File: Choose an Excel File by clicking the button
        2.) 

        

    """
    
    # Title Label
    textbox_label = ttk.Label(content_frame, text=title)
    textbox_label.pack(pady=15)
    
    frame = ttk.Frame(content_frame)
    frame.pack(fill="both", expand=False)
    
    text_area = tk.Text(frame, wrap="none", font=menu_font)
    text_area.pack(padx=15, pady=15, fill="both", expand=True)
    
    # Inserting Text which is read only
    text_area.insert(tk.INSERT, info_text)


def clear_window():
    # Destroy everything except Menu, content frame, and boolian values
    for widget in content_frame.winfo_children():
        if widget != Menu_button and widget != Info_button and widget != Exit_button and widget != content_frame:
            widget.destroy()


#stops program
def exit_program():
    root.destroy()




#Main program config:

root = tk.Tk()

#window header title
root.title("Excel File Parser")



#setting tkinter window size (fullscreen windowed)
root.state('zoomed')

# Set the minimum width and height for the window
root.minsize(400, 300)


menu_font = ("Arial", 14)

# Create a style for the Menu button
menu_button_style = ttk.Style()
menu_button_style.configure("Menu.TButton", font=menu_font, foreground='white', background='#2589bd')

# Create a style for the Info button
info_button_style = ttk.Style()
info_button_style.configure("Toggle.TButton", font=menu_font, foreground='white', background='#5C946E')

# Create a style for the Exit button
exit_button_style = ttk.Style()
exit_button_style.configure("Toggle.TButton", font=menu_font, foreground='white', background='#B3001B')


# Add Main Menu button at the top left of the window
Menu_button = ttk.Button(root, text="Menu", command=display_button, style="Accent.TButton")
Menu_button.place(x=15, y=12)

# Add an Info button
Info_button = ttk.Button(root, text="Info", command=info, style="Info.TButton")
Info_button.place(x=15, y=66)  # Position below Menu_button


# Add an Exit button
Exit_button = ttk.Button(root, text="Exit", command=exit_program, style="Exit.TButton")
Exit_button.place(x=15, y=120)  # Position below Info_button


# Add a darkmode toggle switch
switch = ttk.Checkbutton(text="Light mode", style="Switch.TCheckbutton", command=sv_ttk.toggle_theme)
switch.place(x=15, y=260)  # Position below toggle switch


# Create a content frame for the main content area
content_frame = ttk.Frame(root)
content_frame.place(x=175, y=5, relwidth=.8, relheight=.95)  # Use relative dimensions for expansion


# Bind the window close event to the exit_program function
root.protocol("WM_DELETE_WINDOW", exit_program)



import ctypes as ct
#dark titlebar - ONLY WORKS IN WINDOWS 11!!
def dark_title_bar(window):
    """
    MORE INFO:
    https://learn.microsoft.com/en-us/windows/win32/api/dwmapi/ne-dwmapi-dwmwindowattribute
    """
    window.update()
    DWMWA_USE_IMMERSIVE_DARK_MODE = 20
    set_window_attribute = ct.windll.dwmapi.DwmSetWindowAttribute
    get_parent = ct.windll.user32.GetParent
    hwnd = get_parent(window.winfo_id())
    rendering_policy = DWMWA_USE_IMMERSIVE_DARK_MODE
    value = 2
    value = ct.c_int(value)
    set_window_attribute(hwnd, rendering_policy, ct.byref(value), ct.sizeof(value))


dark_title_bar(root)

#trying to get the taskbar icon to work
myappid = 'Excel.File.Parser.V2' # arbitrary string
ct.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)


# Theme
sv_ttk.set_theme("dark")

# Initial function
display_button()

#this is the loop that keeps the window persistent 
root.mainloop()
