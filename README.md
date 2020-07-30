# file_renamer
A Python-script to easily rename files in a given directory.

This script requires a .xlsx-file named 'file_renamer' in directory 'c:/local/'.
This .xlsx-file requires 2 columns ('A' and 'B') without headers, where column 'A' stores the current filenames and column 'B' stores the required filenames. 

Script execution:
* a path containing the files to rename is requested from the user, using the command prompt.
* each filename in the provided directory is searched in column 'A' of the .xlsx-file and, when found, renamed to the filename provided in column 'B' (in the same row) of the .xlsx-file.
