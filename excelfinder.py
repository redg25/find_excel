from typing import List, Optional
import os
import sys
import zipfile
import xmltodict
import shutil
import argparse
import pandas as pd
from dataclasses import dataclass, field

# Constants set by the sys arguments
SIZE_THRESHOLD = 0
VALUES_TO_SEARCH = []

@dataclass
class ExcelFile:
    id: int
    path: str
    size: int
    are_values_found: bool = False
    file_too_big: bool = False
    sheet_names: List[str] = field(default_factory=list)

def get_sheet_details(file_path: str) -> List[str]:
    # Modified from original script found on stackoverflow
    # https://stackoverflow.com/questions/12250024/how-to-obtain-sheet-names-from-xls-files-without-loading-the-whole-file
    sheets = []
    folder_path = os.path.split(file_path)[0]

    # Make a temporary directory with the file name
    directory_to_extract_to = os.path.join(folder_path,'temp_excelfinder')
    if os.path.isdir(directory_to_extract_to):
        shutil.rmtree(directory_to_extract_to)
    os.mkdir(directory_to_extract_to)

    # Extract the xlsx file as it is just a zip file
    zip_ref = zipfile.ZipFile(file_path, 'r')
    zip_ref.extractall(directory_to_extract_to)
    zip_ref.close()

    # Open the workbook.xml which is very light and only has meta data, get sheets from it
    path_to_workbook = os.path.join(directory_to_extract_to, 'xl', 'workbook.xml')
    with open(path_to_workbook, 'r') as f:
        xml = f.read()
        dictionary = xmltodict.parse(xml)
        for sheet in dictionary['workbook']['sheets']['sheet']:
            sheet_details = {
                'id': sheet['@sheetId'], # can be sheetId without @ for some versions
                'name': sheet['@name'] # can be name without @
            }
            sheets.append(sheet_details['name'])

    # Delete the extracted files directory
    shutil.rmtree(directory_to_extract_to)
    return sheets


def look_for_values_in_file(file: ExcelFile, values_to_search: list) -> bool:
    """
    Takes a dataclass Excel file and a list of string then
    return True if all string are in the excel file
    else returns False
    """
    for sheet in file.sheet_names:
        df = pd.read_excel(file.path, sheet_name=sheet)
        for column in df.values:
            if match_values_in_files(column,values_to_search):
                return True
    return False


def match_values_in_files(file_values,values_to_search):

    def is_substring(col_val, list_of_values):
        for val in list_of_values:
            if col_val.find(val) != -1:
                return True
        return False

    # Get all cell values where the values to search are substrings
    matched_cells = [x for x in file_values if is_substring(str(x),values_to_search)]
    # For each cell, identify which values to search is mapped
    # and remove them form the list of values to search
    for file_value in matched_cells:
        value_to_remove = [x for x in values_to_search if file_value.find(x) != -1]
        for val in value_to_remove:
            values_to_search.remove(val)
    if not values_to_search:
        return True

def are_all_cells_in_file(file: ExcelFile, cells_to_search: list, size_threshold: float) -> bool:
    """
    Given an ExcelFile instance and a list of strings to find in an Excel file
    Returns True if all string(s) are found else returns False
    If the file is too big, user asked to confirm if he wants to search it
    because it takes a long time for the script to open it. If the file is not searched
    the ExcelFile instance cells_not_searched is set to True
    """
    # If the file size is above a threshold ask user to proceed with searc
    if file.size > size_threshold:
        print(f'The size of {file.path} is {file.size} Mo, the search for sheet names was done\n'
              f'but it might take a long time to search for values in cells.')
        still_search = input("If you still want to process this file, type 'y' else press any key:")

        if still_search == 'y':
            # Proceed with search in cells
            return look_for_values_in_file(file, cells_to_search.copy())
        else:
            file.file_too_big = True
            return False
    else:
        # Search in cells
        return look_for_values_in_file(file, cells_to_search.copy())


def find_values(excel_files: ExcelFile,search_cells: bool):
    """
    Search if a set of strings is in the sheet names of a file,
    then search in the cells of this file if requested
    """
    for excel_file in excel_files:
        temp_values_to_search = VALUES_TO_SEARCH.copy()
        excel_file.sheet_names = get_sheet_details(excel_file.path)
        # Search in the sheet names
        match_values_in_files(excel_file.sheet_names,temp_values_to_search)
        # If there are still values to search and the search through the cell is requested
        if temp_values_to_search and search_cells:
            excel_file.are_values_found = are_all_cells_in_file(excel_file, temp_values_to_search, SIZE_THRESHOLD)
        elif not temp_values_to_search:
            excel_file.are_values_found = True


def print_results(excel_files):
    """
    Print results on the terminal.
    Only files with all values found are considered are successful.
    For transparency, it also display files which haven't been search due to size limitation
    """
    def print_found_files(files):
        """Format for a file result"""
        for file in files:
            print(f'\t-id: {file.id}; file: {file.path}\n')

    # Files where all values were found
    found_files = [x for x in excel_files if x.are_values_found]
    # Files for which cells haven't been searched du to size limitation
    files_too_big = [x for x in excel_files if not x.are_values_found and x.file_too_big]
    print('\n\nResult of the search:\n')
    if not found_files and not files_too_big:
        print('No files found...')
    if found_files:
        print(f'The below files contain: {",".join(VALUES_TO_SEARCH)}')
        print_found_files(found_files)
    if files_too_big:
        print('The cells of the below files have not been searched due to file size limitation:')
        print_found_files(files_too_big)


def parse_arguments() -> (float,bool):
    global SIZE_THRESHOLD, VALUES_TO_SEARCH

    def validate_size(size:str) -> float:
        try:
            return float(size)
        except:
            print('Wrong argument input...')
            print(f'--size has to be a float in Mo. Ex: "3.2"')
            sys.exit()

    def validate_values(str_to_search:str) -> Optional[List[str]]:
        if len(str_to_search) > 0:
            return str_to_search.split(',')
        else:
            return None

    parser = argparse.ArgumentParser()
    parser.add_argument('--search',
                        help='sheet names to search; separated by ","',
                        type=validate_values,
                        required=True)
    parser.add_argument('--cells',
                        action='store_true',
                        help='values to search in cells; separated by ","',)
    parser.add_argument('--size',
                        help='max threshold for the file size limitation\n'
                             'User will be asked if he wants to bypass it',
                        type=validate_size,
                        default=1)
    parser.add_argument('--subdir',
                        action='store_true',
                        help='search through all sub-directories')
    args = parser.parse_args()

    SIZE_THRESHOLD = args.size
    VALUES_TO_SEARCH = args.search

    return args.cells, args.subdir


def get_list_of_excel_files_in_directory(subdir: bool) -> List[ExcelFile]:
    list_of_files = []
    directory = os.getcwd()
    for (dir_path, dir_names, filenames) in os.walk(directory):
        filenames = [x for x in filenames if (x.endswith('.xlsx') or x.endswith('.xls')) and not x.startswith('~$')]
        list_of_files += [os.path.join(dir_path, file) for file in filenames]
        if not subdir:
            break
    # Create a list of ExcelFile objects. The file size is converted to Mo.
    list_of_files = [ExcelFile(i, x, round(os.path.getsize(x)/1000000,1)) for i, x in enumerate(list_of_files)]
    return list_of_files


def launch_file(excel_files, id):
    file = [x for x in excel_files if str(x.id) == id]
    if file:
        os.system(f"start EXCEL.EXE {file[0].path}")
        print('File is being launched')
    else:
        print(f'No file with the id: {id}\n'
              f'Exiting script...')


def start_script():
    """Main script steps"""
    search_cells, sub_dir = parse_arguments()
    # Get the list of all files in directory tree at given path
    excel_files = get_list_of_excel_files_in_directory(sub_dir)
    # Look in files if the values to search are in sheet names and/or in cells
    find_values(excel_files,search_cells)
    print_results(excel_files)
    file_id = input('To open one of those file, type its id (Any other input will stop the script): ')
    launch_file(excel_files, file_id)


if __name__ == '__main__':
    start_script()
