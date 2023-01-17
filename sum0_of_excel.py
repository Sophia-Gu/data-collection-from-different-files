import os
import re
import pandas as pd
import openpyxl
from dir_output import file_path_rename


# 已完成给定路径下数据的填写，但是需要改进文件路径的自动读取和明明
def file_name(file_dir):
    for root, dirs, files in os.walk(file_dir):
        return files


# 取出文件中的YF值并规范写法
def rename_YF(xls_dir_path):
    yf_str = re.findall(r"\d+\.?\d*", xls_dir_path)
    yf_value_str = yf_str[0]
    return yf_value_str





# 通过循环迭代可获得excel所在的文件目录
def sum_of_values(xls_dir_path, str='yF='):
    """
    功能： 已实现一级合并，即寻找到模板表格中对应的数据我位置，填入一级汇总表格中
    :param xls_dir_path:表格所在文件目录
    :param str: 第一级汇总文件的开头索引字符串
    :return: 无
    """
    # 表格文件名
    sum_xlsx_name = file_path_rename(xls_dir_path, str)+'_SUM.xlsx'
    # 表格存储完整路径
    sum_xlsx_path = xls_dir_path + '\\' + sum_xlsx_name

    #  (曾经为预防出现xls和xlsx并存的情况)
    # sum_xlsx_path = sum_xls_path + 'x'
    # if os.path.exists(sum_xls_path):
    #     os.remove(sum_xls_path)

    # 如果存在同名表格，即数据已经处理过则将表格删除后重新建立
    if os.path.exists(sum_xlsx_path):
        os.remove(sum_xlsx_path)
    files = file_name(xls_dir_path)
    new_header = ['YH', 'YL', 'R', 'sp.E', 'DM', 'DCH4']
    new_workbook = openpyxl.Workbook()
    work_sheet = new_workbook.active
    work_sheet.append(new_header)
    # 取excel所在文件名为第一个表头
    first_col = ['ZF']  #第一列数据
    # yf_format = rename_YF(xls_dir_path)
    # 循环读取每一excel文件中对应的数值，并存到列表中，每一个文件存为一个行数据
    for i in range(0, len(files)):
        origin_xl_name = files[i]
        # 找到对应的Excel的名字
        if '.xlsx' in origin_xl_name:
            sheet_name = os.path.join(xls_dir_path, origin_xl_name)
            df = pd.read_excel(sheet_name)
            data_YH = float(df.iloc[0, 1])
            data_YL = float(df.iloc[0, 2])
            data_R = float(df.iloc[4, 1])
            data_spE = float(df.iloc[0, 11])
            data_DM = float(df.iloc[0, 7])
            data_DCH4 = float(df.iloc[1, 7])
            new_row = [data_YH, data_YL, data_R, data_spE, data_DM, data_DCH4]
            # 在work_sheet 中将数据整行存入
            work_sheet.append(new_row)
            # 读取文件名中的浓度数据
            col_num_str = re.findall(r"\d+\.?\d*", files[i])
            col_num = float(col_num_str[0])
            first_col.append(col_num)

    # 在表前插入一行（openpyxl的行列从1开始技术），但其他库通常从0开始
    work_sheet.insert_cols(1)
    # 将文件名中的数据写入第一列
    for i in range(1, len(first_col)+1):
        work_sheet.cell(i, 1, first_col[i-1])

    new_workbook.save(filename=sum_xlsx_path)
