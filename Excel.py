from openpyxl import load_workbook
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

try:
    # Get the workbook directory/name
    workbook_directory = input('Enter the directory where the workbook is stored: ')
    # Load the workbook.
    wb = load_workbook(workbook_directory)
except FileNotFoundError:
    print('File \"' + workbook_directory + '\" not found.')
