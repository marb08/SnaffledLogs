import openpyxl
from openpyxl.utils import get_column_letter
import argparse
import glob
import os
import re
import sys

# Define the regular expression pattern to extract the desired information
pattern = r"\](.+) \[(.+)\] {(.+)}<(.+)>\((.+?)\).+?(.*)"

def parse_arguments():
    parser = argparse.ArgumentParser(description='Parse Snaffler log file(s) and save data to XLSX.')
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('-l', '--log_file', help='Path to the log file containing the Snaffler logs')
    group.add_argument('-j', '--json_file', help='Process the json file containing the Snaffled logs')
    group.add_argument('-d', '--directory', help='Process all files containing the Snaffler logs with the specified extension in the current directory')
    parser.add_argument('-x', '--file_extension', help='File extension to filter files when using -a option')
    parser.add_argument('-o', '--output_file', default='snaffler_logs', help='Output file name')

    return parser.parse_args()

def validate_arguments(args, program_name):
    if not (args.log_file or args.directory or args.json_file):
        sys.exit(f"\n{program_name}: error: at least one of the arguments -l/--log_file, -j/--json_file or -d/--directory is required.")

    if args.directory and not args.file_extension:
        sys.exit(f"\n{program_name}: error: when using -d/--directory, you must provide also -x/--file_extension for file extension.")

def set_headers(sheet):
    headers = ['Timestamp', 'Entry Type', 'Triage Color Level', 'Rule Name', 'File Path', 'File Content/File Type']
    for col_num, header in enumerate(headers, 1):
        col_letter = get_column_letter(col_num)
        sheet[f"{col_letter}1"] = header
    
def sanitize_data(data):
    # Define a regular expression to match invalid characters
    invalid_char_pattern = re.compile(r'[\x00-\x08\x0B\x0C\x0E-\x1F]')

    # Replace invalid characters with a placeholder (you can modify this as needed)
    sanitized_data = invalid_char_pattern.sub('_', data)

    return sanitized_data

def process_log_file(file_path, sheet, row_num):
    with open(file_path, 'r') as file:
        for line in file:
            line = line.strip()
            matches = re.findall(pattern, line)
            if matches:
                timestamp, entry_type, triage_level, rule_name, file_path, line_content = matches[0]
                content = re.sub(r'\\r\\n', '', line_content)
                content = re.sub(r'\\r|\\n|\\t', '', content)
                content = re.sub(r'\\\s', ' ', content)

                row_data = [timestamp, entry_type, triage_level, rule_name, file_path, content]
                for col_num, data in enumerate(row_data, 1):
                    col_letter = get_column_letter(col_num)
                    sanitized_data = sanitize_data(data)
                    sheet[f"{col_letter}{row_num}"] = sanitized_data

                row_num += 1

def process_log_files(log_files, sheet, row_num, file_extension=None):
    for log_file_path in log_files:
        if file_extension:
            _, ext = os.path.splitext(log_file_path)
            if ext[1:] == file_extension:  # Exclude the leading dot from the extension
                print(f"Parsing file: {log_file_path}")
                process_log_file(log_file_path, sheet, row_num)
        elif not file_extension:  # Added condition to process all files when -x is not provided
            print(f"Parsing file: {log_file_path}")
            process_log_file(log_file_path, sheet, row_num)

def save_workbook(workbook, output_file_name):
    output_file = f"{output_file_name}"
    workbook.save(output_file)
    return output_file

def print_banner():
    banner = '''
      __           _   _                         
     (_  ._   _. _|_ _|_ |  _   _| |   _   _   _ 
     __) | | (_|  |   |  | (/_ (_| |_ (_) (_| _> 
                                           _|                        
    '''
    print(banner)

def main():

    
    args = parse_arguments()
    validate_arguments(args, sys.argv[0])

    workbook = openpyxl.Workbook()
    sheet = workbook.active

    set_headers(sheet)

    row_num = 2

    if args.log_file:
        log_files = [args.log_file]
    else:
        if os.path.isdir(args.directory):
            log_files = glob.glob(os.path.join(args.directory, '*'))
        else:
            log_files = glob.glob(f"*{args.directory}")

    process_log_files(log_files, sheet, row_num, file_extension=args.file_extension)


    output_file = save_workbook(workbook, args.output_file)

    print_banner()
    print(f"Log data has been successfully extracted and saved to {output_file}.")

if __name__ == "__main__":
    main()