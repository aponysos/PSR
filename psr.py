#! python3.6

import sys
import logging
import xlrd
import csv

LOGGING_FORMAT =        '[%(levelname)7s] %(asctime)s %(msecs)3d ' + \
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

COL_NAMES = [ \
    '项目类型', '项目编码', '项目名称', '专业类别', '项目年份', '项目经理', \
    '可研批复完成时间', '计划采购申请完成时间', '计划合同签订完成时间', '计划订单完成时间', '计划到货完成时间', \
    '采购申请完成时间', '合同签订完成时间', '订单完成时间', '到货完成时间', \
    '设计委托完成时间', '计划设计批复完成时间', '设计批复完成时间', \
    '工程实施发起时间', '计划要求完工时间', '工程实施完成时间', '计划割接上线时间', '割接上线完成时间', '计划投入试运行时间', \
    '初步验收报告完成时间', '计划初步验收时间', '初步验收批复完成时间', '竣工验收报告完成时间', '计划竣工验收时间', '竣工验收批复完成时间', \
    '结算审计申请完成时间', '结算审计报告完成时间', '决算审计申请完成时间', '决算审计报告完成时间', \
    ]

class Project():
    def __init__(self, sheet):
        self.sheet = sheet
        self.rows = []

    def cell_value(self, row, col):
        value = self.sheet.cell(row, col).value
        logging.debug('(%s, %s) %s', row, col, value)
        return value

    def parse(self):
        logging.info('-------- Begin parsing sheet columns %s', self.sheet.name)
        mapper = [0] * len(COL_NAMES)
        for i in range(sheet.ncols):
            col_name = self.cell_value(1, i)
            if col_name in COL_NAMES:
                index = COL_NAMES.index(col_name)
                mapper[index] = i
                logging.debug('found %s -> %s', i, col_name)
        logging.info('-------- End parsing sheet columns %s', sheet.ncols)

        for i, index in enumerate(mapper):
            if index <= 0:
                logging.warn('not found %s', COL_NAMES[i])

        logging.info('-------- Begin reading sheet rows')
        for i in range(2, sheet.nrows):
            row = []
            for j in mapper:
                if j > 0:
                    row.append(self.cell_value(i, j))
                else:
                    row.append('')
            self.rows.append(row)
        logging.info('-------- End reading sheet rows %s', sheet.nrows)

INPUT_FILE_NAME = 'input/input.xlsx'
input = xlrd.open_workbook(INPUT_FILE_NAME)
logging.info('input file %s opened', INPUT_FILE_NAME)

sheet = input.sheets()[0]
p = Project(sheet)
p.parse()

OUTPUT_FILE_NAME = 'output/output.csv'
with open(OUTPUT_FILE_NAME, 'w', newline='') as output:
    logging.info('output file %s opened', OUTPUT_FILE_NAME)
    writer = csv.writer(output)
    logging.info('-------- Begin writing sheet rows')
    for r in p.rows:
        writer.writerow(r)
    logging.info('-------- End writing sheet rows')

logging.info('End PSR logging ...')
