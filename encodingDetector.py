import os
import chardet
from pyth.plugins.rtf15.reader import Rtf15Reader

def detect_encoding(file_path):
    with open(file_path, 'rb') as f:
        rawdata = f.read()
    encoding_result = chardet.detect(rawdata)
    return encoding_result['encoding']

def detect_tables_in_rtf_word_document(file_path):
    try:
        doc = Rtf15Reader.read(open(file_path, 'rb'))
        if doc.content.tables:
            return True
        return False
    except Exception as e:
        print(f"Error processing RTF Word document '{file_path}': {e}")
        return False

def detect_encodings_for_rtf_word_documents_in_folder(folder_path):
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if os.path.isfile(file_path):
            if filename.endswith('.rtf'):
                encoding = detect_encoding(file_path)
                contains_tables = detect_tables_in_rtf_word_document(file_path)
                if contains_tables:
                    print(f"File: {filename}, Encoding: {encoding}")

# Define the folder path
folder_path = r"C:\Users\Joseph\PycharmProjects\NewEmailAutomation\C ACTIONS TO ADD"  # Change this to your folder path

# Check if the folder exists
if os.path.exists(folder_path) and os.path.isdir(folder_path):
    detect_encodings_for_rtf_word_documents_in_folder(folder_path)
else:
    print("Folder not found.")
