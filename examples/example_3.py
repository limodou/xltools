from xltools import Writer

data = [{'items':[
    {'name':'uliweb', 'url':'https://github.com/limodou/uliweb'},
    {'name':'xltools', 'url':'https://github.com/limodou/xltools'},
]}]
Writer('template_3.xlsx', 'sheet1', 'output_1.xlsx', data)
