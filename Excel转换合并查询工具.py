# -*- coding:utf-8 -*-
import contextlib
import copy
import glob
import os
import pandas as pd
import pathlib
import pywintypes
import time
import win32api
import win32com.client as win32
import xlrd
import xlsxwriter
import tkinter as tk
from tkinter import filedialog
import openpyxl
import xlwt
import csv
from tkinter import filedialog, Tk
import sys
from datetime import date, datetime
import tkinter.filedialog


path = os.getcwd()
# 输入目录
inputdir = path
# 输出目录
outputdir = path + "\\out"
if not os.path.exists(outputdir):
    os.mkdir(outputdir)

"""
转换xls功能
"""


def makexls():
    # 三个参数：父目录；所有文件夹名（不含路径）；所有文件名
    all_exce = glob.glob("*.xlsx")
    if (len(all_exce) == 0):
        print("当前目录不存在xlsx文件")
        pass
    else:
        for parent, dirnames, filenames in os.walk(inputdir):
            for fn in filenames:
                if fn.split('.')[-1] == "xlsx":
                    filedir = os.path.join(parent, fn)
                    print("当前进行到:%s" % (filedir))
                    excel = win32.gencache.EnsureDispatch('Excel.Application')
                    wb = excel.Workbooks.Open(filedir)
                    # xlsx: FileFormat=51
                    # xls:  FileFormat=56
                    wb.SaveAs(
                        (os.path.join(outputdir, fn.replace('xlsx', 'xls'))), FileFormat=56)
                    wb.Close()
                    excel.Application.Quit()
        print("转换完成!")
    input("按Enter返回主菜单")


"""
转换xlsx功能
"""


def makexlsx():
    # 三个参数：父目录；所有文件夹名（不含路径）；所有文件名
    all_exce = glob.glob("*.xls")
    if (len(all_exce) == 0):
        print("当前目录不存在xls文件")
        pass
    else:
        for parent, dirnames, filenames in os.walk(inputdir):
            for fn in filenames:
                if fn.split('.')[-1] == "xls":
                    filedir = os.path.join(parent, fn)
                    print("当前进行到:%s" % (filedir))
                    excel = win32.gencache.EnsureDispatch(
                        'Excel.Application')
                    wb = excel.Workbooks.Open(filedir)
                    # xlsx: FileFormat=51
                    # xls:  FileFormat=56
                    wb.SaveAs(
                        (os.path.join(outputdir, fn.replace('xls', 'xlsx'))), FileFormat=51)
                    wb.Close()
                    excel.Application.Quit()
        print("转换完成!")
    input("按Enter返回主菜单")


"""
合并xls功能
"""


# 读取当前路径下面全部的Excel
def pakxls():
    root = tk.Tk()
    root.withdraw()

    # 选择文件夹位置
    filelocation = os.path.normpath(
        filedialog.askdirectory(initialdir=os.getcwd()))
    lst = []

    # 读取文件夹下所有文件（xls和xlsx都读取）
    for i in glob.glob(filelocation + "\\\\" + "*.*"):
        if os.path.splitext(i)[1] in [".xls", ".xlsx"]:
            lst.append(pd.read_excel(i))
    print("选择输出保存文件位置:")
    # 保存合并后的excel文件
    writer = pd.ExcelWriter(filedialog.asksaveasfilename(title="保存", initialdir=filelocation,
                                                         defaultextension="xlsx", filetypes=[("Excel 工作簿", "*.xlsx"), ("Excel 97-2003 工作簿", "*.xls")]))
    pd.concat(lst).to_excel(writer, 'all', index=False)
    writer.save()
    print('\n%d个文件已经合并成功！' % len(lst))


"""
xlsx文件转csv文件
"""
# 生成的csv文件名


def xlsx_to_csv_pd():
    # 实现选择本地文件夹
    path = os.getcwd()
    root = tk.Tk()
    root.withdraw()
    print("选取您需要转换xlsx的文件!")
    time.sleep(1)
    while True:
        print("请选取xlsx文件！")
        file = tkinter.filedialog.askopenfilename()
        if file.split('.')[-1] != "xlsx":
            print("请选取xlsx文件！")
        else:
            data_xls = pd.read_excel(file, index_col=0)
            data_xls.to_csv(path + '\\out\\转换csv.csv', encoding='utf-8')
            break


"""
批量xlsx文件转csv文件
"""
# 生成的csv文件名


def xlsx_to_csv_all():
    # 实现选择本地文件夹
    path = os.getcwd()
    all_exce = glob.glob("*.xlsx")
    if (len(all_exce) == 0):
        print("当前目录不存在xlsx文件")
        pass
    else:
        for parent, dirnames, filenames in os.walk(inputdir):
            for fn in filenames:
                if fn.split('.')[-1] == "xlsx":
                    data_xls = pd.read_excel(fn, index_col=0)
                    data_xls.to_csv(path + "\\out\\" + fn +
                                    ".csv", encoding='utf-8')


"""
批量csv文件转xlsx文件
"""
# 生成的csv文件名


def csv_to_xlsx_all():
    # 实现选择本地文件夹
    path = os.getcwd()
    all_exce = glob.glob("*.csv")
    if (len(all_exce) == 0):
        print("当前目录不存在csv文件")
        pass
    else:
        for parent, dirnames, filenames in os.walk(inputdir):
            for fn in filenames:
                if fn.split('.')[-1] == "csv":
                    csv = pd.read_csv(fn, encoding='utf-8')
                    csv.to_excel(path + "\\out\\" + fn +
                                 ".xlsx", sheet_name='data')


"""
csv文件转换成xlsx文件
"""


def csv_to_xlsx_pd():
    path = os.getcwd()
    root = tk.Tk()
    root.withdraw()
    print("选取您需要转换的csv文件!")
    time.sleep(1)
    file = tkinter.filedialog.askopenfilename()
    while True:
        print("请选取csv文件！")
        file = tkinter.filedialog.askopenfilename()
        if file.split('.')[-1] != "csv":
            print("请选取csv文件！")
        else:
            csv = pd.read_csv(file, encoding='utf-8')
            csv.to_excel(path + '\\out\\转换xlsx.xlsx', sheet_name='data')
            break



    


"""
表格字符串查询工具
"""


def printFinder(val):
    print(val)


def getusefile():
    # 查当前目录下所有xls xlsx文件，返回文件名列表
    usefile = []
    excelfile = sorted(pathlib.Path('.').glob('**/*.xls'))
    usefile = [str(tpfile) for tpfile in excelfile]
    return copy.deepcopy(usefile)


def rdusefile(fileName, checkvalue):
    # 读一个文件，并在文件单元格中查找目标数据，如果找到就返回文件名及数据
    data = xlrd.open_workbook(fileName)  # 打开当前目录下名为 fileName 的文档

    worksheets = data.sheet_names()  # 返回book中所有工作表的名字
    findout = []

    for filenum in range(len(worksheets)):
        # 打开excel文件的第filenum张表
        sheet_1 = data.sheets()[filenum]  # 通过索引顺序获取sheet表
        nrows = sheet_1.nrows  # 获取该sheet中的有效行数
        ncols = sheet_1.ncols  # 获取该sheet中的有效列数
        getdata = []

        # 读取文件数据
        for rowNum in range(0, nrows):
            tep1 = []
            for colNum in range(0, ncols):
                tep1.append(sheet_1.row(rowNum)[colNum].value)
                if checkvalue in str(sheet_1.row(rowNum)[colNum].value):
                    result = []
                    local = fileName.split('.')
                    result.append("文件:"+fileName+" 的表 " +
                                  worksheets[filenum]+" 找到了 ")
                    for cnt in range(0, ncols):
                        result.append(str(sheet_1.row(rowNum)[cnt].value))
                    printFinder(result)

    return copy.deepcopy(findout)


def checkvalue(val):
    # 在当前目录的所有Excel表里找一个字符的位置
    # 获取当前目录内所有Excel 文件列表
    print("开始找  "+val)
    filelist = getusefile()
    check = []
    # 在每一个文件中查找目标数据
    if filelist:
        for filetp in filelist:
            findout = rdusefile(filetp, val)
            if findout:
                check.extend(findout)
    return copy.deepcopy(check)


while True:
    os.system("cls")
    print("====================Excel文件工具箱====================")
    print("请选择需要的功能！请将本程序放到需要转换的文件目录中")
    print("")
    print("1. xlsx批量转换xls文件\n2. xls批量转换xlsx文件\n3. 合并所有xlsx/xls文件\n4. csv文件转换成xlsx文件\n5. xlsx文件转csv文件\n6. xls模糊查询工具\n7. 批量xlsx文件转csv文件\n8. 批量csv文件转xlsx文件\n0. 退出程序")
    print("")
    
    print("当前工作目录:%s" % (path))
    a = int(input("请输入需要转换的格式, 选择序号:\n"))
    if a == 1:
        makexls()
    elif a == 2:
        makexlsx()
    elif a == 3:
        print("请选择需要合并的目录")
        pakxls()
        print("全部合并完成!")
        input("按Enter返回主菜单")

    elif a == 4:
        csv_to_xlsx_pd()
        print("csv文件转xlsx文件结束，输出文件在out/转换xlsx.xlsx ")
        input("按Enter返回主菜单")
    elif a == 5:
        xlsx_to_csv_pd()
        print('xlsx文件转csv文件结束，输出文件在out/转换csv.csv ')
        input("按Enter返回主菜单")
    elif a == 6:
        # 查字符在哪里
        while(1):
            print("\n将要找的文件放在同一个文件夹里哦 =。=")
            findVal = input("请输入要找的字:")
            if findVal != "":
                checkall = checkvalue(findVal)
                print(str(checkall))
                print
    elif a == 7:
        xlsx_to_csv_all()
        print('批量xlsx文件转csv文件结束，输出文件在out目录下 ')
        input("按Enter返回主菜单")
    elif a == 8:
        csv_to_xlsx_all()
        print('批量csv文件转xlsx文件结束，输出文件在out目录下 ')
        input("按Enter返回主菜单")
    elif a == 0:
        print("程序即将推出...")
        time.sleep(2)
        exit()
    else:
        print("输入错误返回主菜单!")
        break
