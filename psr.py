#! python3.6

import sys
import logging
import xlrd
import csv

LOGGING_FORMAT =        '[%(levelname)5s] %(asctime)s %(msecs)3d ' + \
                        '{%(filename)s:%(lineno)4d} %(message)s'
LOGGING_DATE_FORMAT =   '%Y-%m-%d %H:%M:%S'

logging.basicConfig(
    level=logging.INFO,
    format=LOGGING_FORMAT,
    datefmt=LOGGING_DATE_FORMAT,
    filename='psr.log',
    filemode='w')

logging.info('Start PSR logging ...')
logging.info('sys.argv: %s', sys.argv)
logging.info('sys.path: %s', sys.path)

mapper0 = [ \
    1, 2, 3, 4, 5, 31, \
    34, 6, 7, 8, 9, 0, 0, 0, 0, \
    42, 62, 66, \
    75, 76, 78, 82, 84, 10, \
    89, 93, 96, 115, 119, 122, \
    105, 109, 131, 135
    ]

mapper = [(i + 2 if i != 0 else 0) for i in mapper0]

class Project():
    def __init__(self, sheet):
        self.sheet = sheet
        self.rows = []

    def cell_value(self, row, col):
        value = self.sheet.cell(row, col).value
        logging.info('(%s, %s) %s', row, col, value)
        return value

    def parse(self):
        for i in range(2, sheet.nrows):
            row = []
            for j in mapper:
                row.append(self.cell_value(i, j))
            self.rows.append(row)

INPUT_FILE_NAME = 'input/input.xlsx'
input = xlrd.open_workbook(INPUT_FILE_NAME)
logging.info('input file %s opend', INPUT_FILE_NAME)

sheet = input.sheets()[0]
for i in range(1, sheet.ncols):
    logging.info('%s : %s', i, sheet.cell(1, i))

for j in mapper:
    logging.info('%s : %s', j, sheet.cell(1, j))

logging.info('-------- Begin reading sheet %s', sheet.name)
p = Project(sheet)
p.parse()
logging.info('-------- End reading sheet %s', sheet.name)

OUTPUT_FILE_NAME = 'output/output.csv'
with open(OUTPUT_FILE_NAME, 'w', newline='') as output:
    writer = csv.writer(output)
    for r in p.rows:
        writer.writerow(r)

logging.info('End PSR logging ...')
