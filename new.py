import os
import time
import win32com.client
import pythoncom
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Global dictionaries to track renamed files and converted PDFs
renamed_files = {}
converted_pdfs = {}

# Function to convert Word file to PDF
def word_to_pdf(input_file, output_file):
    try:
        # Get the current date in the format mm.dd
        current_date = time.strftime("%m.%d")

        # Get your name
        your_name = "Apurva S"

        # Get the parent folder name
        parent_folder = os.path.basename(os.path.dirname(input_file))

        # Form the new file name with the desired format for the Word file
        new_word_file_name = f"{current_date} - {your_name} - {parent_folder} - CV.docx"

        # Check if the file has already been renamed
        if input_file not in renamed_files:
            # Rename the original Word file to the new Word file name
            time.sleep(0.5)
            os.rename(input_file, os.path.join(os.path.dirname(input_file), new_word_file_name))
            print(f"Renamed {os.path.basename(input_file)} to {new_word_file_name}")

            # Update the renamed_files dictionary
            renamed_files[input_file] = new_word_file_name

        # Initialize the COM library
        pythoncom.CoInitialize()

        # Convert the updated Word file to PDF only if not converted before or if the PDF is outdated
        if input_file not in converted_pdfs or os.path.getmtime(output_file) < os.path.getmtime(input_file):
            word = win32com.client.Dispatch("Word.Application")
            doc = word.Documents.Open(os.path.join(os.path.dirname(input_file), renamed_files[input_file]))

            # Check if the file is not a temporary file (~$) before converting
            if not os.path.basename(input_file).startswith("~$"):
                pdf_output_file = os.path.splitext(os.path.join(os.path.dirname(input_file), renamed_files[input_file]))[0] + ".pdf"

                # Check if the PDF file already exists and delete it
                if os.path.exists(pdf_output_file):
                    os.remove(pdf_output_file)

                # Save the PDF
                doc.SaveAs(pdf_output_file, FileFormat=17)
                print(f"Converted {os.path.basename(renamed_files[input_file])} to PDF: {os.path.basename(pdf_output_file)}")

                # Update the converted_pdfs dictionary
                converted_pdfs[input_file] = pdf_output_file

            doc.Close()
            word.Quit()

        # Uninitialize the COM library
        pythoncom.CoUninitialize()
    except Exception as e:
        # Get the current date and time for the timestamp
        current_time = time.strftime("%Y-%m-%d %H:%M:%S")

        # Update the error message to include the timestamp
        error_msg = f"[{current_time}] Error converting {os.path.basename(input_file)} to PDF: {str(e)}"
        print(error_msg)

# Global variable to hold the observer instance
observer = Observer()

# Watchdog event handler for file system events
class FileHandler(FileSystemEventHandler):
    def is_temp_file(self, file_name):
        return file_name.startswith("~$")

    def on_modified(self, event):
        if not event.is_directory and (event.src_path.endswith('.doc') or event.src_path.endswith('.docx')):
            input_file = event.src_path

            # Check if the file is a temporary file (~$) and skip processing it
            if self.is_temp_file(os.path.basename(input_file)):
                return

            # Add a delay of 1 second before renaming and converting the file
            time.sleep(0.5)
            output_file = os.path.splitext(input_file)[0] + ".pdf"
            word_to_pdf(input_file, output_file)

# Function to start monitoring the folder and its subfolders
def start_folder_monitor(folder_path):
    event_handler = FileHandler()
    observer.schedule(event_handler, path=folder_path, recursive=True)  # Set recursive=True to monitor subfolders
    observer.start()
    print(f"Monitoring folder and subfolders: {folder_path}")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

if __name__ == "__main__":
    folder_to_monitor = r"D:\Google Drive Backup\NU OneDrive\Recruitment\Full time - Job search\Fulltime CV\Monitored Folder"
    start_folder_monitor(folder_to_monitor)
