#!/usr/bin/env python3
# encoding: utf-8
"""
    @author: madgd
    @license: (C) Copyright 2020-2021 madgd. All Rights Reserved.
    @contact: madgdtju@gmail.com
    @software:
    @file: excel_spliter.py
    @time: 2020/9/30 5:05 下午
    @desc: split excel by column value
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)) + "/../")
import xlrd
import xlwt
from utils.utils import titleToNumber, filterByList, getCellValues, writeLine, findColNumByName
from os.path import basename, dirname
import argparse
import time

sep = "$\001$"
def excelSplitBySheet(excelPath, outputPath='', columnLabels="A", headLines=1, sheetNum=1, sheetNameKey="", allSheet=False, *sheetLabels):
    """
    输入excel文件路径，按参数拆分excel表，并返回多个文件对象
    :param excelPath: 输入路径
    :param outputPath: 输出路径
    :param columnLabels: 拆分依据列。默认拆分第一个sheet，拆分依据可为多列, 用英文","隔开, 有顺序
    :param headLines: 跳过表头行数
    :param sheetNum: 第几个sheet
    :param sheetNameKey: 按sheet名称的关键词找拆分表，默认为空。
    :param allSheet: 是否根据第一个sheet的拆分依据，自动找到其他表中相应的列，也进行拆分
    :param sheetLabels: 按传入的labels，同时对不同的sheet进行拆分。生成m*n个表。目前不生效
    :return: 返回成功或失败，msg
    """
    err = ""
    ret = True
    # read excel
    wb = xlrd.open_workbook(filename=excelPath)
    excelName = ".".join(basename(excelPath).split(".")[:-1])
    # print(wb.sheet_names())


    # get target sheet
    sheetNames = wb.sheet_names()
    # default
    targetSheet = wb.sheet_by_index(0)
    # if sheetNum set
    if sheetNum != 1:
        if type(sheetNum) is int and sheetNum > 1 and sheetNum <= len(sheetNames):
            targetSheet = wb.sheet_by_index(sheetNum - 1)
        else:
            err = "sheetNum err"
            return False, err
    # if sheetNameKey set
    if sheetNameKey:
        found = False
        for i in range(len(sheetNames)):
            if sheetNameKey in sheetNames[i]:
                found = True
                targetSheet = wb.sheet_by_index(i)
                break
        if not found:
            err = "sheetNameKey not found"
            return False, err
    # print(targetSheet.name, targetSheet.nrows, targetSheet.ncols)


    # sheet header
    header = []
    for i in range(headLines):
        header.append(getCellValues(targetSheet.row(i)))
    # print(header)


    # split by cols
    # find colnames
    cols2split = columnLabels.split(",")
    cols2splitNum = [titleToNumber(i) - 1 for i in cols2split]
    targetColnames = filterByList(header[-1], cols2splitNum)
    # print(targetColnames)
    # print(findColNumByName(header[-1], targetColnames))
    sheets = [targetSheet]
    # if split all sheet
    if allSheet:
        sheets = wb.sheets()
    else:
        sheetNames = [targetSheet.name]
    rowGroupsBySheet = {}
    headerBySheet = {}
    for sheet in sheets:
        sheetName = sheet.name
        # find cols2splitNum by targetColnames
        tmpHeader = []
        for i in range(headLines):
            tmpHeader.append(getCellValues(sheet.row(i)))
        # print(tmpHeader)
        headerBySheet[sheetName] = tmpHeader
        tmpCols2splitNum = findColNumByName(tmpHeader[-1], targetColnames)
        # print(tmpCols2splitNum)
        #
        allRows = sheet.get_rows()
        for i in range(headLines):
            next(allRows)
        for row in allRows:
            dicKey = sep.join([str(i) for i in getCellValues(filterByList(row, tmpCols2splitNum))])
            if dicKey in rowGroupsBySheet:
                if sheetName in rowGroupsBySheet[dicKey]:
                    rowGroupsBySheet[dicKey][sheetName].append(row)
                else:
                    rowGroupsBySheet[dicKey][sheetName] = [row]
            else:
                rowGroupsBySheet[dicKey] = {sheetName: [row]}


    # save output excels
    if outputPath == "":
        outputPath = dirname(excelPath) + "/%s_split_%s" % (excelName, time.strftime("%Y_%m_%d-%H_%M", time.localtime()))
    if not os.path.exists(outputPath):
        os.mkdir(outputPath)
    for k in rowGroupsBySheet:
        workbook = xlwt.Workbook()
        for sheetName in sheetNames:
            sheet = workbook.add_sheet(sheetName)
            if sheetName not in rowGroupsBySheet[k]:
                rowGroupsBySheet[k][sheetName] = []
            curr = 0
            # header
            for line in headerBySheet[sheetName]:
                writeLine(sheet, line, curr)
                curr += 1
            for line in rowGroupsBySheet[k][sheetName]:
                writeLine(sheet, getCellValues(line), curr)
                curr += 1
        workbook.save("%s/%s_%s.xls" % (outputPath, excelName, k))


def main():
    print("in main")
    # excelSplitBySheet()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='split excel.')
    parser.add_argument('excelPath', metavar='input', type=str,
                        help='input excel')
    parser.add_argument('-o', metavar='output', type=str, default="",
                        help='output path, default input dic')
    parser.add_argument('-c', metavar='columnLabels', type=str, default="A",
                        help='column labels in alphabat to split, multi label sep by ",". default "A"')
    parser.add_argument('-l', metavar='headLines', type=int, default=1,
                        help='header lines, default 1')
    parser.add_argument('-s', metavar='sheetNum', type=int, default=1,
                        help='which sheet to split, default 1')
    parser.add_argument('-k', metavar='sheetNameKey', type=str, default="",
                        help='which sheet to split, search by key. won`t work if not set')
    parser.add_argument('-a', action='store_true',
                        help='if set, split all sheets. default not')
    args = parser.parse_args()

    excelSplitBySheet(args.excelPath, args.o, columnLabels=args.c, headLines=args.l, sheetNum=args.s,\
                      sheetNameKey=args.k ,allSheet=args.a)
    print("split finished!")