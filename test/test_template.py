import os
import sys
sys.path.insert(0, '..')
from xltools import Reader, Writer, Merge, WriteTemplate, Converter
from pprint import pprint
import copy

def test_read_1():
    """
    >>> r = Reader('t1_template.xlsx', 'sheet1', ['t1_data.xlsx'])
    >>> print r.result[0]['items']
    [{u'status': u'S1', u'name': u'A1', u'id': u'ID1'}, {u'status': u'S2', u'name': u'A2', u'id': u'ID2'}, {u'status': u'S4', u'name': u'A4', u'id': u'ID4'}, {u'status': u'S8', u'name': u'A8', u'id': u'ID8'}]
    """

def test_read_2():
    """
    >>> r = Reader('t1_template.xlsx', 'sheet2', ['t1_data.xlsx'])
    >>> print r.result[0]['items']
    [{u'status': u'S1', u'name': u'A1', u'id': u'ID1'}, {u'status': u'S2', u'name': u'A2', u'id': u'ID2'}, {u'status': u'S4', u'name': u'A4', u'id': u'ID4'}, {u'status': u'S8', u'name': u'A8', u'id': u'ID8'}]
    """

def test_read_3():
    """
    >>> r = Reader('t1_template.xlsx', 'sheet3', ['t1_data.xlsx'])
    >>> pprint(r.result)
    [{u'desc': u'DESC1',
      u'id': u'N-ITDM',
      u'name': u'A0902',
      u'request': [{u'field': u'Field1', u'type': u'CHAR(1)'},
                   {u'field': u'Field2', u'type': u'DATE'},
                   {u'field': u'Field3', u'type': u'VARCHAR(2)'}],
      u'response': [{u'field': u'FieldA', u'type': u'CHAR(3)'},
                    {u'field': u'FieldB', u'type': u'DATE'},
                    {u'field': u'FieldC', u'type': u'VARCHAR(5)'}]},
     {u'desc': u'DESC2',
      u'id': u'N-ITSM',
      u'name': u'A0901',
      u'request': [{u'field': u'Field11', u'type': u'CHAR(1)'},
                   {u'field': u'Field21', u'type': u'DATE'},
                   {u'field': u'Field31', u'type': u'VARCHAR(2)'}],
      u'response': [{u'field': u'FieldA1', u'type': u'CHAR(3)'},
                    {u'field': u'FieldB1', u'type': u'DATE'},
                    {u'field': u'FieldC1', u'type': u'VARCHAR(5)'}]},
     {u'desc': u'DESC3',
      u'id': u'FLPM',
      u'name': u'A0903',
      u'request': [{u'field': u'Field12', u'type': u'CHAR(1)'}]}]
    """

def test_read_4():
    """
    >>> r = Reader('t1_template.xlsx', 'sheet1', ['t11_data.xlsx'])
    >>> print r.result[0]['items']
    [{u'status': u'S1', u'name': u'A1', u'id': u'ID1'}, {u'status': u'S2', u'name': u'A2', u'id': u'ID2'}, {u'status': u'S4', u'name': u'A4', u'id': u'ID4'}, {u'status': u'S8', u'name': u'A8', u'id': u'ID8'}]
    """

def test_merge_1():
    """
    >>> a1 = [{u'items': [{u'name': u'A1', u'id': u'ID1'}, {u'name': u'A2', u'id': u'ID2'},
    ...             {u'name': u'A4', u'id': u'ID4'}, {u'name': u'A8', u'id': u'ID8'}],
    ...             u'name': u'Status'}]
    >>> a2 = [{u'items': [{u'name': u'A5', u'id': u'ID1'},
    ...             {u'name': u'A6', u'id': u'ID2'}, {u'name': u'A4', u'id': u'ID4'},
    ...             {u'name': u'A8', u'id': u'ID8'}], u'name': u'Status'}]
    >>> m = Merge()
    >>> #simple combine, just remove duplicated rows
    >>> m.add(a1, a2)
    >>> print m.result
    [{u'items': [{u'name': u'A1', u'id': u'ID1'}, {u'name': u'A2', u'id': u'ID2'}, {u'name': u'A4', u'id': u'ID4'}, {u'name': u'A8', u'id': u'ID8'}, {u'name': u'A5', u'id': u'ID1'}, {u'name': u'A6', u'id': u'ID2'}], u'name': u'Status'}]
    """

def test_merge_2():
    """
    >>> a1 = [{'id':'001', 'items': [{'name': 'A1', 'id': 'ID1'}, {'name': 'A2', 'id': 'ID2'}]},
    ...  {'id':'002', 'items': [{'name': 'A4', 'id': 'ID4'}, {'name': 'A8', 'id': 'ID8'}]},
    ...  ]
    >>> a2 = [{'id':'001', 'items': [{'name': 'A5', 'id': 'ID1'}, {'name': 'A6', 'id': 'ID2'},
    ...                         {'name': 'A6', 'id': 'ID3'}]},
    ...  ]
    >>> m = Merge(keys=['id'], left_join=True, verbose=True)
    >>> m.add(a1, a2)
    >>> print m.result
    [{'items': [{'name': 'A1', 'id': 'ID1'}, {'name': 'A2', 'id': 'ID2'}, {'name': 'A5', 'id': 'ID1'}, {'name': 'A6', 'id': 'ID2'}, {'name': 'A6', 'id': 'ID3'}], 'id': '001'}, {'items': [{'name': 'A4', 'id': 'ID4'}, {'name': 'A8', 'id': 'ID8'}], 'id': '002'}]
    >>> m = Merge(keys=['id', 'items/id'], left_join=True, verbose=True)
    >>> m.add(a1, a2)
    ---- Path= items Skip= {'name': 'A6', 'id': 'ID3'}
    >>> print m.result
    [{'items': [{'name': 'A5', 'id': 'ID1'}, {'name': 'A6', 'id': 'ID2'}], 'id': '001'}, {'items': [{'name': 'A4', 'id': 'ID4'}, {'name': 'A8', 'id': 'ID8'}], 'id': '002'}]
    """

# def test_write_1():
#     """
#     >>> data = [{u'status': u'S1', u'name': u'A1', u'id': u'ID1'}, {u'status': u'S2', u'name': u'A2', u'id': u'ID2'}, {u'status': u'S4', u'name': u'A4', u'id': u'ID4'}, {u'status': u'S8', u'name': u'A8', u'id': u'ID8'}]
#     >>> w = Writer('t1_template.xlsx', 'sheet1', data)
#     """

def test_write_1():
    """
    >>> from openpyxl import load_workbook
    >>> wb = load_workbook('t1_template.xlsx')
    >>> sheet = wb['sheet1']
    >>> t = WriteTemplate(sheet)
    """

def _str(cell, data, _var, _sh):
    return data + ', world'

def _dot(cell, data, n, _var, _sh):
    return '_'*n + data

env = {'str':_str, 'dot':_dot, 'n':4}

class cell(object):
    pass

def test_convert():
    """
    >>> c = Converter("'abc'")
    >>> print c(cell, {}, env)
    abc
    >>> c = Converter('name|str')
    >>> d = {'name':'hello'}
    >>> print c(cell, d, env)
    hello, world
    >>> c = Converter('name|str|dot(2)')
    >>> print c(cell, d, env)
    __hello, world
    >>> c = Converter('name|str|dot(n)')
    >>> print c(cell, d, env)
    ____hello, world
    """

def test_write_2():
    """
    >>> r = Reader('t1_template.xlsx', 'sheet1', ['t1_data.xlsx'])
    >>> w = Writer('t2_template.xlsx', 'sheet1', 't2_output.xlsx', copy.deepcopy(r.result), create=False)
    >>> r1 = Reader('t1_template.xlsx', 'sheet1', ['t2_output.xlsx'])
    >>> r.result == r1.result
    True
    >>> r = Reader('t1_template.xlsx', 'sheet2', ['t1_data.xlsx'])
    >>> w = Writer('t2_template.xlsx', 'sheet2', 't2_output.xlsx', copy.deepcopy(r.result), create=False)
    >>> r1 = Reader('t1_template.xlsx', 'sheet2', ['t2_output.xlsx'])
    >>> r.result == r1.result
    True
    >>> r = Reader('t1_template.xlsx', 'sheet3', ['t1_data.xlsx'])
    >>> w = Writer('t2_template.xlsx', 'sheet3', 't2_output.xlsx', copy.deepcopy(r.result), create=False)
    >>> r1 = Reader('t1_template.xlsx', 'sheet3', ['t2_output.xlsx'])
    >>> r.result == r1.result
    True
    >>> r = Reader('t1_template.xlsx', 'sheet4', ['t1_data.xlsx'])
    >>> w = Writer('t2_template.xlsx', 'sheet4', 't2_output.xlsx', copy.deepcopy(r.result), create=False)
    >>> r1 = Reader('t1_template.xlsx', 'sheet4', ['t2_output.xlsx'])
    >>> r.result == r1.result
    True
    """

def test_write_link():
    """
    >>> w = Writer('t2_template.xlsx', 'sheet5', 't2_output.xlsx', [{'link1':'link1', 'link2':'link2'}], create=False)
    """
