import os
import shutil
from striprtf.striprtf import rtf_to_text

# Constants
SOURCE_FOLDER_PATH = r"C:\Users\Joseph\PycharmProjects\NewEmailAutomation\C ACTIONS TO ADD"
DESTINATION_FOLDER_PATH = r"C:\Users\Joseph\PycharmProjects\NewEmailAutomation\lwo"


def copy_file_to_subfolder(file_path, target_folder_path):
    """Copy a file to a target subfolder."""
    try:
        shutil.copy2(file_path, target_folder_path)  # Using shutil.copy2() to preserve metadata
        print(f"File '{os.path.basename(file_path)}' copied to '{target_folder_path}'")
        return True
    except Exception as e:
        print(f"Error copying file '{os.path.basename(file_path)}': {e}")
        return False


def process_rtf_file(file_path, target_string):
    """Process an RTF file."""
    try:
        # Convert RTF to plain text
        with open(file_path, 'r') as f:
            rtf_content = f.read()
            plain_text = rtf_to_text(rtf_content)
        # Search for the target string in the plain text content
        if target_string in plain_text:
            return True
        else:
            return False
    except Exception as e:
        print(f"Error processing RTF file '{file_path}': {e}")
    return False


def main():
    try:
        # Iterate through destination folders to get target strings
        for folder_name in os.listdir(DESTINATION_FOLDER_PATH):
            folder_path = os.path.join(DESTINATION_FOLDER_PATH, folder_name)
            if os.path.isdir(folder_path):
                target_string = folder_name[:7]  # Extract target string from folder name
                # Iterate through source files to find files containing target string
                for file_name in os.listdir(SOURCE_FOLDER_PATH):
                    file_path = os.path.join(SOURCE_FOLDER_PATH, file_name)
                    if os.path.isfile(file_path):
                        if file_name.endswith('.rtf'):
                            # Process the RTF file
                            if process_rtf_file(file_path, target_string):
                                # Copy the file to the matching folder
                                copy_file_to_subfolder(file_path, folder_path)
                            else:
                                continue
        print("Script execution completed.")
    except Exception as e:
        print(f"Error processing files: {e}")


if __name__ == "__main__":
    main()
