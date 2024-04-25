import os
import shutil

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


def save_attachment_from_file(file_path, target_string):
    """Save attachment from a file to a folder with the matching target string."""
    # Iterate through destination folders
    for folder_name in os.listdir(DESTINATION_FOLDER_PATH):
        folder_path = os.path.join(DESTINATION_FOLDER_PATH, folder_name)
        if os.path.isdir(folder_path) and target_string in folder_name:
            print(f"Attempting to process file '{file_path}' with target string '{target_string}'")
            # Copy the file to the matching folder
            success = copy_file_to_subfolder(file_path, folder_path)
            if not success:
                print(f"Failed to copy file '{file_path}' to '{folder_path}'")


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
                        # Decode the file content using multiple encodings
                        encodings = ['ascii', 'utf-8', 'latin-1', 'utf-16', 'utf-32', 'utf-16-le', 'utf-16-be',
                                     'cp1252', 'cp850', 'cp437']
                        file_content = None
                        for encoding in encodings:
                            try:
                                with open(file_path, 'r', encoding=encoding) as f:
                                    file_content = f.read()
                                break  # If decoding succeeds, break out of the loop
                            except Exception as e:
                                print(f"Error decoding file '{file_name}' with encoding '{encoding}': {e}")
                        if file_content is not None and target_string in file_content:
                            save_attachment_from_file(file_path, target_string)
        print("Script execution completed.")
    except Exception as e:
        print(f"Error processing files: {e}")


if __name__ == "__main__":
    main()
    import os
    import shutil

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


    def save_attachment_from_file(file_path, target_string):
        """Save attachment from a file to a folder with the matching target string."""
        # Iterate through destination folders
        for folder_name in os.listdir(DESTINATION_FOLDER_PATH):
            folder_path = os.path.join(DESTINATION_FOLDER_PATH, folder_name)
            if os.path.isdir(folder_path) and target_string in folder_name:
                print(f"Attempting to process file '{file_path}' with target string '{target_string}'")
                # Copy the file to the matching folder
                success = copy_file_to_subfolder(file_path, folder_path)
                if not success:
                    print(f"Failed to copy file '{file_path}' to '{folder_path}'")


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
                            # Decode the file content using multiple encodings
                            encodings = ['ascii', 'utf-8', 'latin-1', 'utf-16', 'utf-32', 'utf-16-le', 'utf-16-be',
                                         'cp1252', 'cp850', 'cp437']
                            file_content = None
                            for encoding in encodings:
                                try:
                                    with open(file_path, 'r', encoding=encoding) as f:
                                        file_content = f.read()
                                    break  # If decoding succeeds, break out of the loop
                                except Exception as e:
                                    print(f"Error decoding file '{file_name}' with encoding '{encoding}': {e}")
                            if file_content is not None and target_string in file_content:
                                save_attachment_from_file(file_path, target_string)
            print("Script execution completed.")
        except Exception as e:
            print(f"Error processing files: {e}")


    if __name__ == "__main__":
        main()
