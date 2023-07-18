import openpyxl, glob, sys, re
from openpyxl.utils import get_column_letter
import argparse

# Create a new workbook and select the active sheet
workbook = openpyxl.Workbook()
sheet = workbook.active

# Create an argument parser
parser = argparse.ArgumentParser(description='Parse Snaffler log file(s) and save data to XLSX.')
group = parser.add_mutually_exclusive_group(required=True)
group.add_argument('-l', '--log_file', help='Path to the log file containing the Snaffler logs')
group.add_argument('-a', '--extension', help='Process all files containing the Snaffler logs with the specified extension in the current directory')
parser.add_argument('-o', '--output_file', default='snaffler_logs', help='Output file name')

# Get the program name for the error message
program_name = parser.prog

# Parse the command line arguments
args = parser.parse_args()

# Check if at least one of -f or -a is provided
if not (args.log_file or args.log_extension):
    parser.print_usage()
    sys.exit(f"\n{program_name}: error: at least one of the arguments -l/--log_file or -a/--extension is required.")

log_file_path = args.log_file
log_extension = args.extension
output_file_name = args.output_file

# Define the regular expression pattern to extract the desired information
pattern = r"\](.+) \[(.+)\] {(.+)}<(.+)>\((.+?)\).+?(.*)"

# Set the column headers
headers = ['Timestamp', 'Entry Type', 'Triage Level', 'Rule Name', 'File Path', 'File Content']
for col_num, header in enumerate(headers, 1):
    col_letter = get_column_letter(col_num)
    sheet[f"{col_letter}1"] = header

# Initialize the row counter
row_num = 2

# Process log files
if log_file_path:
    log_files = [log_file_path]
else:
    log_files = glob.glob(f"*{log_extension}")

# Process each log file
for log_file_path in log_files:
    with open(log_file_path, 'r') as file:
        # Read the log file line by line and parse each line
        for line in file:
            line = line.strip()  # Remove leading/trailing whitespace
            matches = re.findall(pattern, line)
            if matches:
                timestamp, entry_type, triage_level, rule_name, file_path, line_content = matches[0]
                content=re.sub(r'\\r\\n', '', line_content)
                content=re.sub(r'\\r|\\n|\\t', '', content)
                content=re.sub(r'\\\s', ' ', content)
                
                # Write each match into separate columns
                row_data = [timestamp, entry_type, triage_level, rule_name, file_path, content]
                for col_num, data in enumerate(row_data, 1):
                    col_letter = get_column_letter(col_num)
                    sheet[f"{col_letter}{row_num}"] = data
                
                row_num += 1

# Save the workbook as an XLSX file
output_file = f"{output_file_name}.xlsx"
workbook.save(output_file)

# Print banner in ASCII art
banner = '''
  __           _   _                         
 (_  ._   _. _|_ _|_ |  _   _| |   _   _   _ 
 __) | | (_|  |   |  | (/_ (_| |_ (_) (_| _> 
                                       _|                        
'''
print(banner)
print(f"Log data has been successfully extracted and saved to {output_file}.")