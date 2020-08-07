# file_renamer
A Python-script to easily rename files in a given directory.

This script requires a .xlsx-file with 2 columns (headers are required, or the first file is skipped):
* column 'A' contains the current filename
* column 'B' contains the associated new filename

Script execution:
* the user selects a folder containing the files to rename, using the button "Select folder"
* the user selects a .xlsx-file containing the conversion-matrix, using the button "Select xlsx"
* the user starts the process using the button "Start rename"
* the user exits the application using the button "Quit"
