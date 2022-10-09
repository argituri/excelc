# imports
# from sys import platform as p
import csv
import datetime
import logging
import os
import re
import sys
from pathlib import Path

import pylightxl as xl

logging.basicConfig(stream=sys.stderr, level=logging.DEBUG)
# Check operating system ?


# Read lines from file

filename = re.sub('\s', '_', str(datetime.datetime.now()) + '.csv')
lines = open('lines.txt', 'r').read().splitlines()
xlsxmatch_re = re.compile('.*\.xlsx$')

#excel_file_paths = []

# Add found excel file paths to variable
for line in lines:
    logging.debug('searching .... %s', line)
    # check files in folder, add excel file paths to variable
    try:
        files = os.listdir(line)
        excel_file_paths = [line + '/' + x for x in files if xlsxmatch_re.match(x)]
    except FileNotFoundError as e:
        logging.error("File not found ... %s", e)

if len(excel_file_paths) < 1:
    logging.debug('No excel file paths found!')
    exit()

#try:
for excel_file_path in excel_file_paths:
    logging.debug("opening/writing rows from: %s", excel_file_path)
    db = xl.readxl(fn=excel_file_path)
    with open(filename, 'w+') as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerows(db.ws(ws=db.ws_names[0]).rows)
        # for row in db.ws(ws='Sheet1').rows:
        #    for element in row:
#except Exception as e:
#    logging.error(e)
