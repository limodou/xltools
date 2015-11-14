from xltools import Reader
from pprint import pprint

x = Reader('template_2.xlsx', 'sheet1', 'data_2.xlsx')
pprint(x.result)
