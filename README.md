# SnaffledLogs 🧙‍♂️📝
SnaffledLogs is a simple python script which, using some **magic**, allows to parse output generated by [Snaffler](https://github.com/SnaffCon/Snaffler) into a readable and business xslx file.

## Install and Run
```bash
git clone https://github.com/marb08/SnaffledLogs.git
cd SnaffledLogs/
pip install -r requirements.txt
python3 SnaffledLogs.py
```

## Usage
```bash
  __           _   _                         
 (_  ._   _. _|_ _|_ |  _   _| |   _   _   _ 
 __) | | (_|  |   |  | (/_ (_| |_ (_) (_| _> 
                                       _|                        

Parse Snaffler log file(s) and save data to XLSX.
usage: SnaffledLogs.py [-h] (-l LOG_FILE | -a EXTENSION) [-o OUTPUT_FILE]



optional arguments:
  -h, --help            show this help message and exit
  -l LOG_FILE, --log_file LOG_FILE
                        Path to the log file containing the Snaffler logs
  -a EXTENSION, --extension EXTENSION
                        Process all files with the specified extension in the current directory
  -o OUTPUT_FILE, --output_file OUTPUT_FILE
                        Output file name
```
## Examples
```bash
python3 SnaffledLogs.py -l snaffler_logs.log -o output.xlsx      # Process the logs contained in snaffle_logs.log file.
python3 SnaffledLogs.py -a .log -o output.xlsx                   # Process all the files with .log extension in current directory.
```
## Contributing
Pull requests are welcome.

## License
This script is licensed under the GNU General Public License v3.0.
For more information, please refer to the license text at: https://www.gnu.org/licenses/gpl-3.0.txt
