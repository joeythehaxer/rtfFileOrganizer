import os
import codecs
import chardet

# Constants
TARGET_ENCODING = 'utf-8'
FOLDER_PATH = r'C:\Users\ChrisSingleton\PycharmProjects\Email Automation\.venv\C ACTIONS TO ADD'  # Replace with your folder path

# Function to convert file encoding to UTF-8
def convert_to_utf8(file_path, detected_encoding):
    try:
        # Read the file with the detected encoding
        with open(file_path, 'rb') as f:
            content = f.read()
        # Write the file with UTF-8 encoding
        with codecs.open(file_path, 'w', encoding=TARGET_ENCODING) as f:
            f.write(content.decode(detected_encoding))
        print(f"File '{file_path}' converted to UTF-8.")
    except Exception as e:
        print(f"Error converting file '{file_path}': {e}")

# Function to detect encoding and convert files in the folder
def convert_files_in_folder(folder_path):
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if os.path.isfile(file_path):
            # Detect encoding
            with open(file_path, 'rb') as f:
                detected_encoding = chardet.detect(f.read())['encoding']
            # Convert file encoding to UTF-8
            convert_to_utf8(file_path, detected_encoding)

# Execute the conversion
convert_files_in_folder(FOLDER_PATH)
