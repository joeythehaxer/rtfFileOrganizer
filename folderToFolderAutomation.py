import os
import shutil
from striprtf.striprtf import rtf_to_text

# Constants
SOURCE_FOLDER_PATH = r"C:\Users\Joseph\OneDrive - London Fire Solutions\Documents\GitHub\rtfFileOrganizer\C ACTIONS TO ADD"
DESTINATION_FOLDER_PATH = r"C:\Users\Joseph\OneDrive - London Fire Solutions\Documents\GitHub\rtfFileOrganizer\lwo"


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
    print(target_string)
    """Process an RTF file."""
    # print(target_string)
    # Open the RTF file and read its content
    with open(file_path, 'r') as f:
        rtf_content = f.read()

        # Convert RTF to plain text
        plain_text = repr(rtf_to_text(rtf_content))
        # print(plain_text)
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


def is_folder_empty(folder_path):
    # Check if folder exists
    if not os.path.isdir(folder_path):
        print(f"The path '{folder_path}' is not a valid directory.")
        return None

    # List contents of the folder
    contents = os.listdir(folder_path)

    # Check if the list is empty
    if not contents:
        return True
    else:
        return False


# except Exception as e:
#     # Print an error message if processing fails
#     print(f"Error processing RTF file '{file_path}': {e}")
#     return False


def main():
    try:
        # Iterate through destination folders to get target strings
        folder_names = os.listdir(DESTINATION_FOLDER_PATH)
        folder_names.reverse()
        for folder_name in folder_names:
            folder_path = os.path.join(DESTINATION_FOLDER_PATH, folder_name)
            # print(folder_path)
            if os.path.isdir(folder_path) and is_folder_empty(folder_path):
                # Extract target string from folder name
                target_string = folder_name[:8]
                processed_target_string = ""
                for a in target_string:
                    if a.isalnum():
                        processed_target_string += a
                print("searching for job number " + processed_target_string + " in " + SOURCE_FOLDER_PATH)
                # Iterate through source files to find files containing target string
                file_names = os.listdir(SOURCE_FOLDER_PATH)
                # file_names.reverse()
                for file_name in file_names:

                    file_path = os.path.join(SOURCE_FOLDER_PATH, file_name)
                    if os.path.isfile(file_path):
                        # print(file_path)

                        # Process the RTF file

                        if process_rtf_file(file_path, processed_target_string):
                            # Copy the file to the matching folder
                            copy_file_to_subfolder(file_path, folder_path)
                        else:
                            continue
            else:
                continue
        # Print a message when script execution is completed
        print("Script execution completed.")
    except Exception as e:
        # Print an error message if an exception occurs during processing
        print(f"Error processing files: {e}")


if __name__ == "__main__":
    main()
