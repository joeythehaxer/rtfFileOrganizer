import os
import openpyxl
from striprtf.striprtf import rtf_to_text

# Constants
FOLDER_PATH = r"C:\Users\josep\PycharmProjects\rtfFileOrganizer\C ACTIONS TO ADD"
EXCEL_FILE_PATH = r"C:\Users\josep\PycharmProjects\rtfFileOrganizer\New Microsoft Excel Worksheet.xlsx"
SHEET_NAME = "Sheet1"
CELLS_TO_TARGET = [(1, 1), (2, 2), (3, 3)]  # Example: Targeting cells A1, B2, and C3


def load_excel_workbook(excel_file_path):
    return openpyxl.load_workbook(excel_file_path)


def find_last_empty_row(ws):
    return ws.max_row + 1


def extract_info_from_rtf(file_path):

    extracted_info = []
    with open(file_path, 'r') as f:
        rtf_content = f.read()
        plain_text = rtf_to_text(rtf_content)
        lines = plain_text.split("|")
        print(lines)
        # client string
        clientraw = lines[0].split(">")[1]
        client = ""
        for char in clientraw:
            if char.isalnum() or char == " ":
                client+=char
        # Client Derivative String
        clientDerivative = lines[50]
        # LFS Job String
        jobString = "3197"
        ponumberString = lines[2]
        print(ponumberString)


    return extracted_info


def paste_info_to_excel(ws, extracted_info, last_row):
    for i, cell_text in enumerate(extracted_info):
        ws.cell(row=last_row, column=i + 1).value = cell_text
    return last_row + 1


def save_excel_workbook(wb, excel_file_path):
    wb.save(excel_file_path)


def import_info_from_rtf_folder(folder_path, excel_file_path, sheet_name):
    # Load Excel workbook
    wb = load_excel_workbook(excel_file_path)
    ws = wb[sheet_name]

    # Loop through each RTF file in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith(".rtf"):
            # Construct the full file path
            file_path = os.path.join(folder_path, filename)

            # Extract information from RTF file
            extracted_info = extract_info_from_rtf(file_path)

            # Find the last empty row in Excel sheet
            last_row = find_last_empty_row(ws)

            # Paste the extracted information into a new row in Excel
            paste_info_to_excel(ws, extracted_info, last_row)

    # Save changes to Excel workbook
    save_excel_workbook(wb, excel_file_path)


# Usage
import_info_from_rtf_folder(FOLDER_PATH, EXCEL_FILE_PATH, SHEET_NAME)
