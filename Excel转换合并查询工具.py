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

biao_tou = []
# 获取要合并的所有exce表格


def get_exce():
    all_exce = glob.glob("*.xls")
    print("该目录下有" + str(len(all_exce)) + "个xls表格文件：")
    if (len(all_exce) == 0):
        print("当前目录不存在xls文件")
        input("按Enter返回主菜单")
        pass
    else:
        for i in range(len(all_exce)):
            print(all_exce[i])
        return all_exce
# 打开Exce文件


def open_exce(name):
    fh = xlrd.open_workbook(name)
    return fh
# 获取exce文件下的所有sheet


def get_sheet(fh):
    sheets = fh.sheets()
    return sheets
# 获取sheet下有多少行数据


def get_sheetrow_num(sheet):
    return sheet.nrows
# 获取sheet下的数据


def get_sheet_data(sheet, row, biao_tou_num):
    for i in range(row):
        if (i < biao_tou_num):
            global biao_tou
            values = sheet.row_values(i)
            biao_tou.append(values)
            continue
        values = sheet.row_values(i)
        all_data1.append(values)
    return all_data1
# 获取表头数量


def get_biao_tou_num(exce1, exce2):
    fh = open_exce(exce1)
    fhx = open_exce(exce2)
    sheet_1 = fh.sheet_by_index(0)
    sheet_2 = fhx.sheet_by_index(0)
    row_sum_1 = sheet_1.nrows
    row_sum_2 = sheet_2.nrows
    # 获取第一张sheet表对象有效行数
    # 获取sheet表某一行所有数据类型及值
    for i in range(row_sum_1):
        # 获取sheet表对象某一行数据值
        if (i+1 == row_sum_2):
            return i
        #row_0_value = sheet_1.row_values(0)
        row_content_1 = sheet_1.row_values(i)
        row_content_2 = sheet_2.row_values(i)
        if(row_content_1 == row_content_2):
            continue
        else:
            return i


"""
合并xlsx功能
"""


def merge():
    # 批量表所在文件夹路径
    path = os.getcwd()
    all_exce = glob.glob("*.xlsx")
    outfile = path + '\\out\\汇总.xlsx'
    print("该目录下有" + str(len(all_exce)) + "个xlsx表格文件")
    if (len(all_exce) == 0):
        print("当前目录不存在xlsx文件")
        input("按Enter返回主菜单")
        pass
    else:
        arr = []
        print("开始合并xlsx...")
        if os.path.exists(outfile) == True:
            os.remove(outfile)
            print("清理旧汇总xlsx文件成功")
        else:
            pass
        open(outfile, "w")
        time.sleep(3)
        for parent, dirnames, filenames in os.walk(inputdir):
            for fn in filenames:
                if fn.split('.')[-1] == "xlsx":
                    arr.append(pd.read_excel(fn))
                    # 目标文件的路径
                    writer = pd.ExcelWriter(outfile)
                    pd.concat(arr).to_excel(writer, 'sheet1', index=False)
                    writer.save()
        print("汇总.xlsx输出成功！")
        input("按Enter返回主菜单")
        exit()


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
    print("============Excel文件工具箱============")
    print("请选择需要的功能！请将本程序放到需要转换的文件目录中")
    print("1. xlsx转换成xls\n2. xls转换成xlsx\n3. 合并所有xls\n4. 合并所有xlsx\n5. xls字符串查询工具\n6. 退出程序")
    a = int(input("请输入需要转换的格式, 选择序号:\n"))
    if a == 1:
        makexls()
    elif a == 2:
        makexlsx()
    elif a == 3:
        print("使用本程序只需要把程序放到需要合并表格同目录下")
        all_exce = get_exce()
        # 得到要合并的所有exce表格数据
        if (all_exce == 0):
            print("该目录下无.xls文件！请把程序移动到要合并的表格同目录下！")
            pass
        if (len(all_exce) == 1):
            print("该目录下只有一个.xls文件！无需合并")
            pass
        # 表头数
        print("自动检测表头中......")
        biao_tou_num = get_biao_tou_num(all_exce[0], all_exce[1])
        print("表头数为:", biao_tou_num,)
        guess = input("y/n?")
        if(guess == "n"):
            biao_tou_num = input("请输入表头数:")
            biao_tou_num = int(biao_tou_num)
        all_data1 = []
        # 用于保存合并的所有行的数据
        # 下面开始文件数据的获取
        for exce in all_exce:
            fh = open_exce(exce)
            # 打开文件
            sheets = get_sheet(fh)
            # 获取文件下的sheet数量
            for sheet in range(len(sheets)):
                row = get_sheetrow_num(sheets[sheet])
                # 获取一个sheet下的所有的数据的行数
                all_data2 = get_sheet_data(sheets[sheet], row, biao_tou_num)
                # 获取一个sheet下的所有行的数据
        for i in range(biao_tou_num):
            all_data2.insert(i, biao_tou[i])
        # 表头写入
        new_name = input("清输入新表的名称:")
        # 下面开始文件数据的写入
        new_exce = path + "\\out\\" +new_name+".xls"
        # 新建的exce文件名字
        fh1 = xlsxwriter.Workbook(new_exce)
        # 新建一个exce表
        new_sheet = fh1.add_worksheet()
        # 新建一个sheet表
        for i in range(len(all_data2)):
            for j in range(len(all_data2[i])):
                c = all_data2[i][j]
                new_sheet.write(i, j, c)
        fh1.close()
        # 关闭该exce表
        print("文件合并成功,请查看"+new_exce+"文件！")
        input("按Enter返回主菜单")
    elif a == 4:
        merge()
    elif a == 5:
        # 查字符在哪里
        while(1):
            print("\n将要找的文件放在同一个文件夹里哦 =。=")
            findVal = input("请输入要找的字:")
            if findVal != "":
                checkall = checkvalue(findVal)
                print(str(checkall))
                print
    elif a == 6:
        print("程序即将推出...")
        time.sleep(3)
        exit()            
    else:
        print("输入错误程序即将退出!")
        time.sleep(3)
        exit()
