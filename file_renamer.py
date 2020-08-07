# file_renamer: file_renamer.py
"""This application is a script to rename the files in a single folder,
in a single action.
The application requires a .xlsx-file with 2 columns, where column 'A' contains
the current filename and column 'B' contains the required filename.
"""
# Import statement(s).
from os import path
import sys
from tkinter import Tk
from _c_filerenamerui import FileRenamerUI

def resource_path(relative_path):
    """Get the absolute path to resource."""
    base_path = getattr(sys, '_MEIPASS', path.dirname(path.abspath(__file__)))
    return path.join(base_path, relative_path)

def file_renamer():
    """This is the main function for the file_renamer.
    This function stores the revision information and calls the GUI-class.
    """
    # Variable(s).
    rev_numb = 1.0
    rev_date = "2020-08-07"
    title = f"file_renamer by S. Cassier, rev. {rev_numb} ({rev_date})"
    image_path = resource_path("red_ibis_64x64.ico")
    # Call and loop the GUI.
    root = Tk()
    root.iconbitmap(image_path)
    root.title(title)
    root.resizable(False, False)
    FileRenamerUI(root, rev_numb, rev_date)
    root.mainloop()

if __name__ == "__main__":
    file_renamer()