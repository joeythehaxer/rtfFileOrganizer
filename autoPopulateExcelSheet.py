import os
import re
from openpyxl import load_workbook
from striprtf import striprtf
from langdetect import detect

# Constants
FOLDER_PATH = "C ACTIONS TO ADD"
EXCEL_FILE_PATH = "New Microsoft Excel Worksheet.xlsx"
SHEET_NAME = "Sheet1"


def remove_non_english(text):
    """
    Remove non-English words from the given text.

    Args:
    text (str): Input text containing words.

    Returns:
    str: Text with only English words.
    """
    # Tokenize the text into words
    words = re.findall(r'\b\w+\b', text)

    # Filter out non-English words
    english_words = [word for word in words if any(c.isalpha() for c in word) and word != "FDI"]

    # Join the English words into a sentence
    english_sentence = ' '.join(english_words)

    return english_sentence


def summarize(text):
    """
    Summarize the text by removing non-English words.

    Args:
    text (str): Input text.

    Returns:
    str: Summarized text with only English words.
    """
    # Remove non-English text
    english_text = remove_non_english(text)

    # Return the entire English text
    return english_text


def load_excel_workbook():
    """
    Load the Excel workbook from the specified path.

    Returns:
    openpyxl.workbook.Workbook: Loaded Excel workbook object.
    """
    return load_workbook(EXCEL_FILE_PATH)


def find_last_empty_row(ws):
    """
    Find the last empty row in the specified worksheet.

    Args:
    ws (openpyxl.worksheet.Worksheet): Worksheet object.

    Returns:
    int: Index of the last empty row.
    """
    return ws.max_row + 1


def extract_info_from_rtf(file_path):
    """
    Extract information from the RTF file.

    Args:
    file_path (str): Path to the RTF file.

    Returns:
    list: Extracted information.
    """
    with open(file_path, 'r') as f:
        rtf_content = f.read()
        plain_text = striprtf.rtf_to_text(rtf_content)

    lines = plain_text.split("|")
    client_raw = lines[0].split(">")[1]
    client = "".join(char for char in client_raw if char.isalnum() or char == " ")
    client_derivative = lines[50]
    job_string = "3197"
    po_number_string = lines[2]
    address_string = lines[23]
    postcode_pattern = re.compile(r'\b[A-Z]{1,2}\d{1,2}\s\d[A-Z]{2}\b')
    risk_string = ""

    try:
        postcode_string = postcode_pattern.findall(address_string)[0]
    except IndexError or UnboundLocalError:
        postcode_string = "None"

    tel_string = "NOT SUPPLIED"
    if lines[16]:
        tel_string = lines[16]

    date_raised_string = lines[38]
    target_date_string = lines[46]
    tender_number_pattern = re.compile(r'\b\d{5,7}\b')
    tender_number_list = tender_number_pattern.findall(lines[61])

    try:
        tender_number_string = tender_number_list[0]
    except IndexError:
        tender_number_string = "None"

    raw_action_description_string = lines[61]
    english_action_string = summarize(raw_action_description_string)

    if len(tender_number_list) <= 1:
        extracted_info = [client_derivative, job_string, po_number_string, tender_number_string,
                          date_raised_string, target_date_string, tel_string, address_string,
                          postcode_string, risk_string, english_action_string]
    else:
        extracted_info = [client_derivative, job_string, po_number_string, tender_number_list,
                          date_raised_string, target_date_string, tel_string, address_string,
                          postcode_string, risk_string, english_action_string]

    return extracted_info


def paste_info_to_excel(ws, extracted_info, last_row):
    """
    Paste the extracted information into the Excel worksheet.

    Args:
    ws (openpyxl.worksheet.Worksheet): Worksheet object.
    extracted_info (list): Extracted information to paste.
    last_row (int): Index of the last row in the worksheet.

    Returns:
    int: Updated index of the last row.
    """
    if isinstance(extracted_info[3], list):
        for tender_number in extracted_info[3]:
            extracted_info[3] = tender_number
            for column, cell_text in enumerate(extracted_info, start=2):
                ws.cell(row=last_row, column=column).value = cell_text
            last_row += 1
    else:
        for column, cell_text in enumerate(extracted_info, start=2):
            ws.cell(row=last_row, column=column).value = cell_text
        last_row += 1

    return last_row


def save_excel_workbook(wb):
    """
    Save the Excel workbook to the specified path.

    Args:
    wb (openpyxl.workbook.Workbook): Workbook object to save.
    """
    wb.save(EXCEL_FILE_PATH)


def import_info_from_rtf_folder():
    """
    Import information from RTF files in the specified folder and paste it into an Excel worksheet.
    """
    wb = load_excel_workbook()
    ws = wb[SHEET_NAME]

    for filename in os.listdir(FOLDER_PATH):
        if filename.endswith(".rtf"):
            file_path = os.path.join(FOLDER_PATH, filename)
            extracted_info = extract_info_from_rtf(file_path)
            last_row = find_last_empty_row(ws)
            last_row = paste_info_to_excel(ws, extracted_info, last_row)

    save_excel_workbook(wb)


# Usage
import_info_from_rtf_folder()
