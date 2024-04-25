import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk

# Test values
TEST_EXCEL_FILE = r"C:\Users\Joseph\PycharmProjects\NewEmailAutomation\NEW Project Hub Priorities .xlsx"
TEST_SHEET_NAME = "DIP Tracker NEW "
TEST_OUTPUT_DIR = r"C:\Users\Joseph\PycharmProjects\NewEmailAutomation\lwo"
TEST_ID_CHARACTER_NUM = 7
TEST_IDENTIFYING_NUMBER = "PO number / Action Number"
TEST_RELATED_DATA_2 = "ITEM LOCATION / ADDRESS"
TEST_RELATED_DATA_3 = "POSTCODE"
TEST_IDENTIFYING_NUMBER_COL_NUMBER = 4
TEST_RELATED_DATA_2_COL_NUMBER = 9
TEST_RELATED_DATA_3_COL_NUMBER = 10

# Flag for test mode
TEST_MODE = False


def create_folders():
    global EXCEL_FILE, SHEET_NAME, OUTPUT_DIR, ID_CHARACTER_NUM, IDENTIFYING_NUMBER, RELATED_DATA_2, RELATED_DATA_3, COLUMN_NAMES

    EXCEL_FILE = excel_file_entry.get() if not TEST_MODE else TEST_EXCEL_FILE
    SHEET_NAME = sheet_name_entry.get() if not TEST_MODE else TEST_SHEET_NAME
    OUTPUT_DIR = output_dir_entry.get() if not TEST_MODE else TEST_OUTPUT_DIR
    ID_CHARACTER_NUM = int(id_character_num_entry.get() if not TEST_MODE else TEST_ID_CHARACTER_NUM)
    IDENTIFYING_NUMBER = identifying_number_entry.get() if not TEST_MODE else TEST_IDENTIFYING_NUMBER
    RELATED_DATA_2 = related_data_2_entry.get() if not TEST_MODE else TEST_RELATED_DATA_2
    RELATED_DATA_3 = related_data_3_entry.get() if not TEST_MODE else TEST_RELATED_DATA_3
    COLUMN_NAMES = [IDENTIFYING_NUMBER, RELATED_DATA_2, RELATED_DATA_3]
    IDENTIFYING_NUMBER_COL_NUMBER = int(
        identifying_number_col_number_entry.get() if not TEST_MODE else TEST_IDENTIFYING_NUMBER_COL_NUMBER)
    RELATED_DATA_2_COL_NUMBER = int(
        related_data_2_col_number_entry.get() if not TEST_MODE else TEST_RELATED_DATA_2_COL_NUMBER)
    RELATED_DATA_3_COL_NUMBER = int(
        related_data_3_col_number_entry.get() if not TEST_MODE else TEST_RELATED_DATA_3_COL_NUMBER)
    COLUMN_NUMBERS = [IDENTIFYING_NUMBER_COL_NUMBER, RELATED_DATA_2_COL_NUMBER, RELATED_DATA_3_COL_NUMBER]

    try:
        # Read the Excel file
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, header=None,
                           usecols=COLUMN_NUMBERS,
                           names=COLUMN_NAMES, skiprows=1)

        # Create the output directory if it doesn't exist
        if not os.path.exists(OUTPUT_DIR):
            os.makedirs(OUTPUT_DIR)

        # Get a list of existing folder names in the output directory


        # Iterate through the values in the specified columns and create folder name
        # Iterate through the values in the specified columns and create folder name
        for index, row in df.iterrows():
            existing_folders = os.listdir(OUTPUT_DIR)
            po_number = str(row[COLUMN_NAMES[0]])
            address = str(row[COLUMN_NAMES[1]])
            postcode = str(row[COLUMN_NAMES[2]])

            folder_name = f"{po_number} {address}".replace("/", "-")
            folder_name = "".join(x for x in folder_name if x.isalnum() or x in [' ', '_', "-"])
            if len(folder_name) > 120:
                folder_name = f"{po_number} {postcode}".replace("/", "-")
                folder_name = "".join(x for x in folder_name if x.isalnum() or x in [' ', '_', "-"])

            # Check if any existing folder contains the folder name
            folder_exists = any(
                folder_name[:TEST_ID_CHARACTER_NUM] in existing_folder for existing_folder in existing_folders)
            print(folder_name[:TEST_ID_CHARACTER_NUM])
            # Modify the folder name to make it unique if it already exists
            if folder_exists:
                # new_folder_name = folder_name
                # count = 1
                # while os.path.exists(os.path.join(OUTPUT_DIR, new_folder_name)):
                #     new_folder_name = f"{folder_name}_{count}"
                #     count += 1
                # folder_path = os.path.join(OUTPUT_DIR, new_folder_name)
                # folder_name = new_folder_name
                continue
            else:
                folder_path = os.path.join(OUTPUT_DIR, folder_name)
                os.makedirs(folder_path)
            status_label.configure(text=f"Folder '{folder_name}' created successfully.", background="#008000")

        status_label.config(text="Folder creation complete.", foreground="blue")

    except FileNotFoundError:
        status_label.config(text="Excel file not found.", foreground="red")
    except pd.errors.ParserError:
        status_label.config(text=f"Error reading data from sheet '{SHEET_NAME}' in Excel file.", foreground="red")


def browse_excel_file():
    excel_file_entry.delete(0, tk.END)
    excel_file_entry.insert(0, filedialog.askopenfilename())


def browse_output_dir():
    output_dir_entry.delete(0, tk.END)
    output_dir_entry.insert(0, filedialog.askdirectory())


def toggle_test_mode():
    global TEST_MODE
    TEST_MODE = not TEST_MODE
    if TEST_MODE:
        set_test_values()
    else:
        clear_input_fields()


def set_test_values():
    excel_file_entry.insert(0, TEST_EXCEL_FILE)
    sheet_name_entry.insert(0, TEST_SHEET_NAME)
    output_dir_entry.insert(0, TEST_OUTPUT_DIR)
    id_character_num_entry.insert(0, TEST_ID_CHARACTER_NUM)
    identifying_number_entry.insert(0, TEST_IDENTIFYING_NUMBER)
    related_data_2_entry.insert(0, TEST_RELATED_DATA_2)
    related_data_3_entry.insert(0, TEST_RELATED_DATA_3)
    identifying_number_col_number_entry.insert(0, TEST_IDENTIFYING_NUMBER_COL_NUMBER)
    related_data_2_col_number_entry.insert(0, TEST_RELATED_DATA_2_COL_NUMBER)
    related_data_3_col_number_entry.insert(0, TEST_RELATED_DATA_3_COL_NUMBER)


def clear_input_fields():
    excel_file_entry.delete(0, tk.END)
    sheet_name_entry.delete(0, tk.END)
    output_dir_entry.delete(0, tk.END)
    id_character_num_entry.delete(0, tk.END)
    identifying_number_entry.delete(0, tk.END)
    related_data_2_entry.delete(0, tk.END)
    related_data_3_entry.delete(0, tk.END)
    identifying_number_col_number_entry.delete(0, tk.END)
    related_data_2_col_number_entry.delete(0, tk.END)
    related_data_3_col_number_entry.delete(0, tk.END)


# Create main window
root = tk.Tk()
root.title("Folder Creation GUI")

# Create and place input widgets
label_font = ('Arial', 12)
entry_font = ('Arial', 12)

excel_file_label = ttk.Label(root, text="Excel File Path:", font=label_font)
excel_file_label.grid(row=0, column=0, padx=5, pady=5)
excel_file_entry = ttk.Entry(root, font=entry_font)
excel_file_entry.grid(row=0, column=1, padx=5, pady=5)
excel_file_button = ttk.Button(root, text="Browse", command=browse_excel_file)
excel_file_button.grid(row=0, column=2, padx=5, pady=5)

sheet_name_label = ttk.Label(root, text="Sheet Name:", font=label_font)
sheet_name_label.grid(row=1, column=0, padx=5, pady=5)
sheet_name_entry = ttk.Entry(root, font=entry_font)
sheet_name_entry.grid(row=1, column=1, padx=5, pady=5)

output_dir_label = ttk.Label(root, text="Output Directory:", font=label_font)
output_dir_label.grid(row=2, column=0, padx=5, pady=5)
output_dir_entry = ttk.Entry(root, font=entry_font)
output_dir_entry.grid(row=2, column=1, padx=5, pady=5)
output_dir_button = ttk.Button(root, text="Browse", command=browse_output_dir)
output_dir_button.grid(row=2, column=2, padx=5, pady=5)

id_character_num_label = ttk.Label(root, text="ID Character Length:", font=label_font)
id_character_num_label.grid(row=3, column=0, padx=5, pady=5)
id_character_num_entry = ttk.Entry(root, font=entry_font)
id_character_num_entry.grid(row=3, column=1, padx=5, pady=5)

identifying_number_label = ttk.Label(root, text="Column Name for Identifying Number:", font=label_font)
identifying_number_label.grid(row=4, column=0, padx=5, pady=5)
identifying_number_entry = ttk.Entry(root, font=entry_font)
identifying_number_entry.grid(row=4, column=1, padx=5, pady=5)

related_data_2_label = ttk.Label(root, text="Column Name for Related Data 2:", font=label_font)
related_data_2_label.grid(row=5, column=0, padx=5, pady=5)
related_data_2_entry = ttk.Entry(root, font=entry_font)
related_data_2_entry.grid(row=5, column=1, padx=5, pady=5)

related_data_3_label = ttk.Label(root, text="Column Name for Related Data 3:", font=label_font)
related_data_3_label.grid(row=6, column=0, padx=5, pady=5)
related_data_3_entry = ttk.Entry(root, font=entry_font)
related_data_3_entry.grid(row=6, column=1, padx=5, pady=5)

identifying_number_col_number_label = ttk.Label(root, text="Column Number for Identifying Number:", font=label_font)
identifying_number_col_number_label.grid(row=7, column=0, padx=5, pady=5)
identifying_number_col_number_entry = ttk.Entry(root, font=entry_font)
identifying_number_col_number_entry.grid(row=7, column=1, padx=5, pady=5)

related_data_2_col_number_label = ttk.Label(root, text="Column Number for Related Data 2:", font=label_font)
related_data_2_col_number_label.grid(row=8, column=0, padx=5, pady=5)
related_data_2_col_number_entry = ttk.Entry(root, font=entry_font)
related_data_2_col_number_entry.grid(row=8, column=1, padx=5, pady=5)

related_data_3_col_number_label = ttk.Label(root, text="Column Number for Related Data 3:", font=label_font)
related_data_3_col_number_label.grid(row=9, column=0, padx=5, pady=5)
related_data_3_col_number_entry = ttk.Entry(root, font=entry_font)
related_data_3_col_number_entry.grid(row=9, column=1, padx=5, pady=5)

create_folders_button = ttk.Button(root, text="Create Folders", command=create_folders)
create_folders_button.grid(row=10, columnspan=3, padx=5, pady=10)

status_label = ttk.Label(root, text="", font=label_font)
status_label.grid(row=11, columnspan=3, padx=5, pady=5)

test_mode_button = ttk.Button(root, text="Toggle Test Mode", command=toggle_test_mode)
test_mode_button.grid(row=12, columnspan=3, padx=5, pady=10)

root.mainloop()