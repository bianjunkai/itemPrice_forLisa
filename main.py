#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Date    : 2017-03-16
# @Function: As Lisa requires, read the datasheet.xls as a reference,then check the line of target excel and put the price inside
# @Author  : Junkai BIAN

import os
import string
import xlrd
import xlwt
from xlutils.copy import copy

BASEDIR = os.path.abspath(os.path.dirname(__file__))
EXCELFILE = os.path.join(BASEDIR, 'datasheet.xls')
DATAFILE = os.path.join(BASEDIR, 'data.xls')

"""
Return the data in the datasheet
"""


def get_reference():

    reference = None
    if os.path.isfile(EXCELFILE):
        reference = xlrd.open_workbook(EXCELFILE)
    print("Datasheet is :%s"%EXCELFILE)
    return reference

"""
Return the price of target
"""


def get_price(ref, target):
    reference = ref
    target = target
    table = reference.sheet_by_index(0)
    nrows = table.nrows
    maletargetlist = table.col_values(1)
    malepricelist = table.col_values(2)
    femaletagetlist = table.col_values(3)
    femalepricelist = table.col_values(4)
    for i in range(nrows):
        if target == maletargetlist[i]:
            return malepricelist[i]
        elif target == femaletagetlist[i]:
            return femalepricelist[i]
        else:
            continue


def get_data_file():
    file = DATAFILE
    filepath = input('请将xls文件路径粘贴进去，如果程序里已经指定了文件则按Enter键继续')
    is_valid = False            # 验证文件
    try:
        filepath = [file, filepath][filepath != '']
        print(filepath)
        # 判断给出的路径是不是xls格式
        if os.path.isfile(filepath):
            filename = os.path.basename(filepath)
            if filename.split('.')[1] == 'xls':
                is_valid = True
        data = None
        if is_valid:
            data = xlrd.open_workbook(filepath,formatting_info=True)
            global DATAFILE
            DATAFILE = filepath
    except Exception as e:
        print('你操作错误：%s' % e)
        return None
    return data


def get_data(datafile, ref):
    data_ref = datafile
    reference = ref
    table_ref = data_ref.sheet_by_index(1)
    nrows = table_ref.nrows
    file_write = copy(datafile)
    table_write= file_write.get_sheet(1)
    componentlist = table_ref.col_values(1)
    for i in range(nrows):
        if componentlist[i] != "":
            price = get_price(reference,componentlist[i])
            if price is None:
                continue
            print("部件编号:%s\t"%componentlist[i])
            print("价格：%d\n"%price)
            table_write.write(i,3,price)
            print("已写入在D,%d单元格\n"%(i+1))
    file_write.save(DATAFILE)
    return "******处理完毕************"


def main():
    reference = get_reference()
    data = get_data_file()
    get_data(data,reference)

if __name__ == '__main__':
    main()
