import os
import hashlib

# Constants
CHUNK_SIZE = 1024

def is_file_corrupted(file_path):
    try:
        # Method 1: Check file size
        file_size = os.path.getsize(file_path)
        if file_size == 0:
            return True  # File is likely corrupted (empty)

        # Method 2: Calculate file hash (MD5 checksum)
        md5_hash = hashlib.md5()
        with open(file_path, 'rb') as file:
            while True:
                chunk = file.read(CHUNK_SIZE)
                if not chunk:
                    break
                md5_hash.update(chunk)
        calculated_hash = md5_hash.hexdigest()

        # You might have a pre-calculated hash stored somewhere for comparison
        # stored_hash = "previously_calculated_hash"
        # if calculated_hash != stored_hash:
        #     return True  # File is corrupted

        # Method 3: Try opening the file (if it's a known file format)
        # For example, if it's a text file
        # with open(file_path, 'r') as file:
        #     file_content = file.read()

        return False  # File is not corrupted

    except Exception:
        return True  # File is likely corrupted

def check_corrupted_files(folder_path):
    # Get a list of all files in the folder
    files = os.listdir(folder_path)

    # Iterate through each file
    for file_name in files:
        file_path = os.path.join(folder_path, file_name)
        if is_file_corrupted(file_path):
            print(f"The file '{file_name}' is corrupted.")
        else:
            print(f"The file '{file_name}' is not corrupted.")

# Specify the path to the subfolder
SUBFOLDER_PATH = r"C:\Users\Joseph\PycharmProjects\NewEmailAutomation\C ACTIONS TO ADD"

# Call the function to check for corrupted files
check_corrupted_files(SUBFOLDER_PATH)
