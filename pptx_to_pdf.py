'''
Converts all pptx files under a folder into pdf files.
It requires PowerPoint and it is limited to Windows.

Useful information regarding Python and PowerPoint at:
https://stackoverflow.com/questions/45316940/error-using-comtypes-client-to-convert-a-pptx-file-into-pdf
https://stackoverflow.com/questions/64783044/export-powerpoint-to-mp4-by-script-with-dynamic-slide-duration
If you don't want to use Python but VBA:
https://www.brightcarbon.com/blog/how-to-use-vba-in-powerpoint/

Aldebaro. Mar 2024.
'''
import os
import comtypes.client
import argparse
from pathlib import Path


def create_folder_if_not_exists(folder_path):
    """
    Creates a folder if it does not exist.

    Args:
        folder_path (str): The path to the folder.

    Returns:
        None
    """
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        print(f"Folder '{folder_path}' created successfully.")
    else:
        print(f"Folder '{folder_path}' already exists.")


def PPTtoPDF(inputFileName, outputFileName):
    formatType = 32
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    if outputFileName[-3:] != 'pdf':
        outputFileName = outputFileName + ".pdf"
    deck = powerpoint.Presentations.Open(inputFileName)
    deck.SaveAs(outputFileName, formatType)  # formatType = 32 for ppt to pdf
    deck.Close()
    powerpoint.Quit()


def get_files_with_extension(path, recurse=True, file_types=(".pptx")):
    file_count = 0
    iterator = os.walk(path) if recurse else ((next(os.walk(path))), )
    list_of_files = list()
    for root, dirs, file_names in iterator:
        for file_name in file_names:
            list_of_files.append(os.path.join(root, file_name)) if file_name.lower().endswith(
                file_types) else 0
    return list_of_files


def convert_pptx_to_pdf(input_folder, output_folder, recurse=True):
    list_of_files = get_files_with_extension(input_folder, recurse=recurse)
    N = len(list_of_files)
    if recurse:
        print("Found", N, "files with extension pptx with a recursive search inside folder", input_folder)
    else:
        print("Found", N, "files with extension pptx under folder",
              input_folder, "(did not conduct a recursive search)")
    file_num = 1
    for input_file in list_of_files:
        # input_file = os.path.join(root, filename)
        filename = os.path.basename(input_file)
        output_file = os.path.splitext(filename)[0] + ".pdf"
        outputFileName = os.path.join(output_folder, output_file)
        # we convert to absolute paths to make sure PowerPoint finds the files
        inputFileName = str(Path(input_file).resolve())
        outputFileName = str(Path(outputFileName).resolve())
        print("#", file_num, ": input =", inputFileName,
              "=>", "output =", outputFileName)
        PPTtoPDF(inputFileName, outputFileName)
        file_num += 1


if __name__ == "__main__":
    # Create an ArgumentParser instance
    parser = argparse.ArgumentParser(
        description="Convert all pptx files in a folder into pdf files.",
        epilog="Usage: <input_folder> <output_folder>")

    # Add positional arguments
    parser.add_argument("input_folder", type=str,
                        help="Input folder with PPTX files.")
    parser.add_argument("output_folder", type=str,
                        help="Output folder where the PDF files will be saved (it will be created if it does not exist).")

    # Add optional arguments
    parser.add_argument('-r', '--recursively', action='store_true',
                        help="Search input folder recursively (look inside subfolders).")

    # Parse the command-line arguments
    args = parser.parse_args()

    # Access the parsed arguments
    input_folder_path = args.input_folder
    output_folder_path = args.output_folder
    recurse = args.recursively

    # Convert to PDF
    create_folder_if_not_exists(output_folder_path)
    convert_pptx_to_pdf(input_folder_path, output_folder_path, recurse=recurse)
