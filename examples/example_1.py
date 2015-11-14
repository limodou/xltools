from xltools import Reader
from pprint import pprint

x = Reader('template_1.xlsx', 'sheet1', 'data_1.xlsx')
pprint(x.result)
