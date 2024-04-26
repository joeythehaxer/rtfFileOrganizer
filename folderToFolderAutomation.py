import os
import shutil
from striprtf.striprtf import rtf_to_text

# Constants
SOURCE_FOLDER_PATH = r"C:\Users\Joseph\PycharmProjects\NewEmailAutomation\C ACTIONS TO ADD"
DESTINATION_FOLDER_PATH = r"C:\Users\Joseph\PycharmProjects\NewEmailAutomation\lwo"


def copy_file_to_subfolder(file_path, target_folder_path):
    """Copy a file to a target subfolder."""
    try:
        # Copy the file to the target folder
        shutil.copy(file_path, target_folder_path)  # Using shutil.copy2() to preserve metadata
        # Print a success message
        print(f"File '{os.path.basename(file_path)}' copied to '{target_folder_path}'")
        return True
    except Exception as e:
        # Print an error message if copying fails
        print(f"Error copying file '{os.path.basename(file_path)}': {e}")
        return False


def process_rtf_file(file_path, target_string):
    """Process an RTF file."""
    # print(target_string)
    # Open the RTF file and read its content
    with open(file_path, 'r') as f:
        rtf_content = f.read()

        # Convert RTF to plain text
        plain_text = repr(rtf_to_text(rtf_content))
        # print(repr(plain_text))
    # Search for the target string in the plain text content
        if target_string in plain_text:
            # Return True if the target string is found
            print("string found")
            return True
        else:
            # Return False if the target string is not found
            print("target string wasn't found")
            return False

# except Exception as e:
#     # Print an error message if processing fails
#     print(f"Error processing RTF file '{file_path}': {e}")
#     return False


def main():
    try:
        # Iterate through destination folders to get target strings
        for folder_name in os.listdir(DESTINATION_FOLDER_PATH):
            folder_path = os.path.join(DESTINATION_FOLDER_PATH, folder_name)
            # print(folder_path)
            if os.path.isdir(folder_path):
                # Extract target string from folder name
                target_string = folder_name[:8]
                processed_target_string = ""
                for a in target_string:
                    if a.isalnum():
                        processed_target_string += a
                print("searching for job number " + processed_target_string + " in " + SOURCE_FOLDER_PATH)
                # Iterate through source files to find files containing target string
                for file_name in os.listdir(SOURCE_FOLDER_PATH):

                    file_path = os.path.join(SOURCE_FOLDER_PATH, file_name)
                    if os.path.isfile(file_path):
                        # print(file_path)

                        # Process the RTF file

                        if process_rtf_file(file_path, processed_target_string):
                            # Copy the file to the matching folder
                            copy_file_to_subfolder(file_path, folder_path)
                        else:
                            continue

        # Print a message when script execution is completed
        print("Script execution completed.")
    except Exception as e:
        # Print an error message if an exception occurs during processing
        print(f"Error processing files: {e}")


if __name__ == "__main__":
    main()
