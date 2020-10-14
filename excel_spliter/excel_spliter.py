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
import xlrd

def excelSplitBySheet(excel_path, output_path, columnLabels="A", headLines=1, sheetNamePre="", allSheet=False, *sheetLabels):
    """
    输入excel文件路径，按参数拆分excel表，并返回多个文件对象
    :param excel_path: 输入路径
    :param output_path: 输出路径
    :param columnLabels: 拆分依据列。默认拆分第一个sheet，拆分依据可为多列, 用英文","隔开
    :param headLines: 跳过表头行数
    :param sheetNamePre: 按sheet名称的前缀找拆分表，默认为空。
    :param allSheet: 是否根据第一个sheet的拆分依据，自动找到其他表中相应的列，也进行拆分
    :param sheetLabels: 按传入的labels，同时对不同的sheet进行拆分。生成m*n个表。目前不生效
    :return: 返回成功或失败，msg
    """
    # 读取excel
    wb = xlrd.open_workbook(filename=excel_path)
    print(wb.sheet_names())
    # 存储excel


def main():
    print("in main")
    # excelSplitBySheet()


if __name__ == '__main__':
    excel_path = ""
    output_path = ""
    excelSplitBySheet(excel_path, output_path)