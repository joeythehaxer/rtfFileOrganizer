import os
import subprocess

def convert_rtf_to_docx(input_folder, output_folder, pandoc_path):
    # Ensure output folder exists
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Get list of RTF files in input folder
    rtf_files = [f for f in os.listdir(input_folder) if f.endswith('.rtf')]

    # Convert each RTF file to DOCX
    for rtf_file in rtf_files:
        input_file_path = os.path.join(input_folder, rtf_file)
        output_file_path = os.path.join(output_folder, os.path.splitext(rtf_file)[0] + '.docx')
        subprocess.run([pandoc_path, input_file_path, '-o', output_file_path])

if __name__ == "__main__":
    input_folder = 'input_folder'  # Specify the folder containing RTF files
    output_folder = 'output_folder'  # Specify the folder to save the converted DOCX files
    pandoc_path = r"C:\Users\Joseph\AppData\Local\Pandoc\pandoc.exe"  # Specify the full path to the Pandoc executable

    convert_rtf_to_docx(input_folder, output_folder, pandoc_path)
