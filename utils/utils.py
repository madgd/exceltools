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
    # xlrd ctype_text
    return [i.value if i.ctype not in [0, 5, 6] else "" for i in cols]

def writeLine(sheet, line, row=0, startCol=0):
    for column, heading in enumerate(line, startCol):
        sheet.write(row, column, heading)

def findColNumByName(row, targetColnames):
    """

    :param row:
    :param targetColnames:
    :return:
    """
    return [row.index(name) for name in targetColnames]

def checkEmptyLine(cols):
    """

    :param line:
    :return:
    """
    for i in cols:
        if i.ctype not in [0, 5, 6]:
            return False
    return True