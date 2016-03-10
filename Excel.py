from openpyxl import load_workbook, Workbook
import os


# Variables
user_documents = os.path.expanduser('~\Documents')  # Users documents dir.
save_directory = user_documents + '\Excel in Python'  # Save the default save dir.

# Check if the default save location exists
if not os.path.exists(save_directory):
    # Make the directory
    os.mkdir(save_directory)
    # Inform the user
    print('[INFO] Program save directory created: ' + save_directory)

# Check if the workbook save directory exists.
if not os.path.exists(save_directory + '\Saved Workbook Directories.txt'):
    # Create a txt file to hold save directories
    saved_workbooks = open(save_directory + '\Saved Workbook Directories.txt', 'a')
    # Inform the user
    print('[INFO] Workbook save directory created: ' + save_directory + '\Saved Workbook Directories.txt')

# Prompt the user on whether or not to use an existing workbook.
use_existing = input('Would you like to use an existing workbook (y/n): ')

while True:
    try:
        if 'y' in use_existing:
            # Get the workbook directory/name
            workbook_directory = input('Enter the directory where the workbook is stored: ')

            try:
                # Load the workbook.
                wb = load_workbook(workbook_directory)
            except:
                # Get the filename and file_extension of the given file.
                filename, file_extension = os.path.splitext(workbook_directory)
                # Print the issue
                print('[ERROR] Incorrect file type \'' + file_extension + '\'.')
                # Continue
                continue
        else:
            # Prompt the user for the filename.
            filename = input('Enter the name of the new workbook: ')
            # Create the new workbook
            wb = Workbook(filename)

    except FileNotFoundError:
        print('[INFO] File \"' + workbook_directory + '\" not found.')
        continue
    break

