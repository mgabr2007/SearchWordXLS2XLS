# Install required packages
# Run these commands in your terminal or command prompt:
# pip install xlwt xlrd xlutils
# pip install PySimpleGUI

import os
import xlrd
from xlutils.copy import copy
import PySimpleGUI as sg

# Create GUI to browse for folder
folder_path = sg.popup_get_folder("Select folder to search in:")

# Create GUI to browse for Excel file containing search words
sg.theme('DarkAmber')
layout = [[sg.Text('Select Excel file containing search words:')],
          [sg.Input(), sg.FileBrowse(file_types=(("Excel Files", "*.xls"), ("All Files", "*.*")))],
          [sg.Submit(), sg.Cancel()]]
window = sg.Window('File Search', layout)

# Read search words from selected Excel file
if window.Read()[0] == 'Submit':
    file_path = window.Element(0).Get()
    workbook = xlrd.open_workbook(file_path)
    worksheet = workbook.sheet_by_index(0)
    search_words = [str(cell.value) for cell in worksheet.col(0)]
    window.Close()

    # Create output Excel file
    output_file_path = sg.popup_get_file("Select output Excel file:", file_types=(("Excel Files", "*.xls"), ("All Files", "*.*")))
    output_workbook = xlwt.Workbook()
    output_worksheet = output_workbook.add_sheet("Results")
    output_worksheet.write(0, 0, "Search Word")
    output_worksheet.write(0, 1, "File Path")

    # Search for files in selected folder
    row = 1
    for search_word in search_words:
        file_paths = []
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if search_word in file:
                    file_paths.append(os.path.join(root, file))

        # Write results to output Excel sheet
        if file_paths:
            for file_path in file_paths:
                output_worksheet.write(row, 0, search_word)
                output_worksheet.write(row, 1, file_path)
                row += 1

    # Save output Excel file
    output_workbook.save(output_file_path)

    # Display success message
    sg.popup("File search completed!")
else:
    window.Close()
