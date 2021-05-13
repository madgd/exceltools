#!/usr/bin/env python
# encoding: utf-8
"""
    @author: madgd
    @license: (C) Copyright 2020-2021 madgd. All Rights Reserved.
    @contact: madgdtju@gmail.com
    @software: 
    @file: utils.py
    @time: 2020/10/15 3:06 下午
    @desc: some common tools
"""
from copy import copy

def titleToNumber(s: str) -> int:
    # start from 1
    return sum( (ord(a) - 64) * (26 ** i)  for i, a in enumerate(s[::-1]))

def filterByList(array, keys):
    """
    filter target array by keys
    :param array:
    :param keys:
    :return:
    """
    tmp = []
    for key in keys:
        tmp.append(array[key])
    return tmp

def getCellValues(cols):
    """

    :param cols:
    :return:
    """
    # openpyxl.cell data_type
    # cell.value can be None
    return [cell.value if cell.data_type in ['s', 'f', 'n', 'inlineStr', 'str'] and cell.value else "" for cell in cols]

def copyLine(sheet, line, row=0, startCol=0, styles=False):
    """
    copy one line with styles
    :param sheet:
    :param line:
    :param row:
    :param startCol:
    :return:
    """
    # row or col start with 1 in openpyxl
    for col in range(len(line)):
        cell = line[col]
        # print(row, col, cell.value)
        new_cell = sheet.cell(row=row+1, column=col+startCol+1, value=cell.value)
        if styles and cell.has_style:
            new_cell.font = copy(cell.font)
            new_cell.border = copy(cell.border)
            new_cell.fill = copy(cell.fill)
            new_cell.number_format = copy(cell.number_format)
            new_cell.protection = copy(cell.protection)
            new_cell.alignment = copy(cell.alignment)
            new_cell.comment = copy(cell.comment)
            new_cell.hyperlink = copy(cell.hyperlink)


def findColNumByName(row, targetColNames):
    """

    :param row:
    :param targetColnames:
    :return:
    """
    value_row = [cell.value for cell in row]
    return [value_row.index(cell.value) for cell in targetColNames]

def checkEmptyLine(cols):
    """

    :param line:
    :return:
    """
    for i in cols:
        if i.data_type and i.data_type in ['s', 'f', 'n', 'inlineStr', 'str']:
            return False
    return True