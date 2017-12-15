#! python3.6

import logging
import sys

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
logging.info('End PSR logging ...')
