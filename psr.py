#!

import logging
import xlrd
import xlwt

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

input = xlrd.open_workbook('input/input.xlsx')
output = xlrd.open_workbook('output/output.xlsx')
table = input.sheets()[0]
cell_A1 = table.cell(1,1).value
logging.info(cell_A1)

table2 = output.sheets()[0]
cell_A3 = table2.cell(3,0).value
logging.info(cell_A3)
# 类型 0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
ctype = 1
value = cell_A1 + '*'
logging.info(value)
xf = 0 # 扩展的格式化
table2.put_cell(3, 1, ctype, value, xf)

workbook = xlwt.Workbook(encoding='utf-8')
data_sheet = workbook.add_sheet('demo')
data_sheet.write(0, 0, value)
workbook.save('temp.xls')
