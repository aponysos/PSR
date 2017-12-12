#! PYTHON36

import logging
import xlrd
import csv

LOGGING_FORMAT =        '[%(levelname)5s] %(asctime)s:%(msecs)3d ' + \
                        '{%(filename)s:%(lineno)4d} %(message)s'
LOGGING_DATE_FORMAT =   '%Y-%m-%d %H:%M:%S'

logging.basicConfig(
    level=logging.INFO,
    format=LOGGING_FORMAT,
    datefmt=LOGGING_DATE_FORMAT,
    filename='psr_csv.log',
    filemode='w')

class Project():
    def __init__(self, sheet):
        self.sheet = sheet

    def cell_value(self, row, col):
        value = self.sheet.cell(row, col).value
        logging.info(value)
        return value

    def parse(self): 
        self.project_no = self.cell_value(1, 1)
        self.project_name = self.cell_value(1, 2)
        self.mananger = self.cell_value(1, 3)
        self.project_type = self.cell_value(1, 4)
        self.limit = self.cell_value(1, 5)
        self.project_rop = self.cell_value(1, 6)
        self.starttime = self.cell_value(5, 5)

    def dump(self):
        csv

logging.info('Start PSR_CSV logging ...')

INPUT_FILE_NAME = 'input/input.xlsx'
input = xlrd.open_workbook(INPUT_FILE_NAME)
logging.info('input file %s opend', INPUT_FILE_NAME)

projects = []
for s in input.sheets():
    logging.info('Begin reading sheet %s', s.name)
    p = Project(s)
    p.parse()
    projects.append(p)
    logging.info('End reading sheet %s', s.name)

OUTPUT_FILE_NAME = 'output/output.csv'
with open(OUTPUT_FILE_NAME, 'w', newline='') as output:
    writer = csv.writer(output)
    for p in projects:
        writer.writerow([p.project_no, p.project_type, p.project_name, \
            p.mananger, p.limit, p.project_rop])
        

logging.info('Stop PSR_CSV logging ...')
