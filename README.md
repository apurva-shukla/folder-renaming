# folder-renaming
## Intro
Applying to jobs requires adapting a main CV to each job. This script eliminates the manual process of renaming your Word (aka, *.doc or *.docx file) and exporting it to a *.pdf file.

script.py scans a folder for any new Word(*.doc or *.docx file), and when it detects a new file, it renames it in the following format:
MM.DD - [your_name] - [parent_folder_name] - CV

It also then exports the Word file into a PDF file. 

There are additional checks to ensure that the file exporting only happens once, and the .pdf file is always updated. 

It saves me at least 20+ mins everyday doing a very manual error-prone process.

## How to run
1. Download script.py to your local computer and run in it in IDE. Ensure that python and pip are installed on your computer already.
2. Edit script.py to identify which folder to monitor (Line 109)
3. Edit script.py to change the file naming format (Line 25)
4. Run the script and create a new folder for each job you apply to 
