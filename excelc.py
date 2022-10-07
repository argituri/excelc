# imports
# from sys import platform as p
import pylightxl as xl
import sys, os, re, csv, datetime, logging
from pathlib import Path
logging.basicConfig(stream=sys.stderr, level=logging.DEBUG)
# Check operating system ?


# Read lines from file

filename = re.sub('', str(datetime.datetime.now()))
lines = Path('lines.txt').read_text()
excel_file_paths = []

# Add found excel file paths to variable
for line in lines:
    logging.debug('searching .... %s', line)
    # check files in folder, add excel file paths to variable
    files = os.listdir(line)
    excel_file_paths.extend(re.findall('.*\.xlsx$', files))

if len(excel_file_paths) < 1:
    logging.debug('No excel file paths found!')
    exit()

try:
    for excel_file_path in excel_file_paths:
        logging("opening file %s", excel_file_path)
        db = xl.readxl(fn=excel_file_path)
        with open('%s.csv', filename) as f:
            writer = csv.writer(f, delimiter=';')
            writer
            writer.writerows(db.ws(ws='Sheet1').rows)
            #for row in db.ws(ws='Sheet1').rows:
            #    for element in row:
except Exception as e:
    logging.error(e);







