# -*- coding: utf-8 -*-
# @Author : PuYang
# @Time   : 2022/10/28
# @File   : main.py
import time

import pandas as pd
from opts import *
import xlwt
import xlrd

global xls_file
xls_file = xlwt.Workbook(encoding='utf-8')

# 第一问
# 获取excel基本信息
def get_zero_data(path):
    data = pd.read_excel(path, sheet_name=sheet_basic_info)
    test_sys_name = data[system_name_str][0]  # 来自表基本信息sheet中C列
    sub_sys = data[sheet_name_str][0]    # 来自表基本信息sheet中B列
    sys_name = "{}{}".format(test_sys_name, database_info_str)
    gen_user = "{}".format(get_owner(input_path))
    create_date = get_current_create_time()
    ncount = len(data[subsys_tag_str])+3
    sheet_count = "{}".format(ncount)
    df = {
        "sys_name":sys_name,
        "gen_user":gen_user,
        "create_date":create_date,
        "sheet_count":sheet_count,
        "sub_sys":sub_sys
    }
    return df

# 封面
def write_sheet_zero(dict, path):
    workbook = xls_file
    worksheet = workbook.add_sheet(cover_str)
    style = xlwt.XFStyle()

    # 设定字体样式
    font0 = xlwt.Font()
    font0.name = 'Arial'
    font0.height = 10 * 20

    font = xlwt.Font()
    font.name = '宋体'
    font.height = 28*20

    font01 = xlwt.Font()
    font01.name = 'Arial'
    font01.height = 28 * 20

    font16 = xlwt.Font()
    font16.name = '宋体'
    font16.height = 16 * 20

    fontA16 = xlwt.Font()
    fontA16.name = 'Arial'
    fontA16.height = 16 * 20

    # 对齐方式的设置(alignment)
    alignment = xlwt.Alignment()
    alignment.vert = 0x00  # 0x00 上端对齐；0x01 居中对齐(垂直方向上)；0x02 底端对齐
    alignment.horz = 0x01  # 0x01 左端对齐；0x02 居中对齐(水平方向上)；0x03 右端对齐
    alignment.wrap = 1  # 自动换行
    style.alignment = alignment

    first_col = worksheet.col(0)  # 第一列
    first_col.width = 996*20  # 第一列列宽
    tall_style = xlwt.easyxf('font:height 5710')  # 设置行高
    first_row = worksheet.row(0)  # 获取sheet页的第一行
    first_row.set_style(tall_style)  # 给第一行设置tall_style样式，也就是行高

    seg0 = ("\n\n\n\n", font0)
    seg1 = ("{}".format(dict['sys_name']), font)
    seg2 = ("({})".format(dict['sub_sys']), font01)
    seg3 = ("\n\n\n\n\n\n", font)
    segf4 = ("  ", fontA16)
    seg40 = (generate_user_str, font16)
    seg41 = ("{}\n  ".format(dict['gen_user']), fontA16)
    seg50 = (generate_date_str, font16)
    seg51 = ("{}\n  ".format(dict['create_date']), fontA16)
    seg60 = (sheet_counts_str, font16)
    seg61 = ("{}\n  ".format(dict['sheet_count']), fontA16)

    worksheet.write_rich_text(0, 0, [seg0, seg1, seg2, seg3, segf4, seg40, seg41, seg50, seg51, seg60, seg61], style)
    workbook.save(path)

# 第二问
# 拷贝基本信息表
def copy_basic_sheet():

    xls = xlrd.open_workbook(input_path)
    new_workbook = xls_file
    table = xls.sheet_by_index(0)
    rows = table.nrows
    cols = table.ncols
    worksheet = new_workbook.add_sheet(sheet_basic_info)
    headstyle = pend_sheet_head_style()
    normalstyle = pend_sheet_normal_style()
    for i in range(0, rows):
        for j in range(0, cols):
            if j == 5 and i != 0 :
                # 添加超链接
                style = xlwt.XFStyle()
                font = style.font
                font.name = "宋体"
                font.height = 10*20
                font.underline = 0x01
                font.colour_index = 0x0C
                borders = xlwt.Borders()
                borders.left = xlwt.Borders.THIN
                borders.right = xlwt.Borders.THIN
                borders.top = xlwt.Borders.THIN
                borders.bottom = xlwt.Borders.THIN
                style.font = font
                style.borders = borders
                link = 'HYPERLINK("#{}!A1";"{}")'.format(table.cell_value(i, j), table.cell_value(i, j))
                worksheet.write(i, j, xlwt.Formula(link), style)
            else:
                if i == 0:
                    worksheet.write(i, j, table.cell_value(i, j), headstyle)
                else:
                    worksheet.write(i, j, table.cell_value(i, j), normalstyle)
    for k in range(0, cols):
        worksheet.col(k).width = 180 * 43
        tall_style = xlwt.easyxf('font:height 256')  # 设置行高
        first_row = worksheet.row(k)
        first_row.set_style(tall_style)
    new_workbook.save(output_path_final)

# 第三问 第四问
# 添加Sheet页
# sheet_name :
# T_TST_SYS_PARK_PEND
# T_TST_SYS_AUDIT_INFO
# T_TST_SYS_AUDT_CHCK
# T_TST_SYS_ADUT_CHCK
def add_all_three_sheets():
    xls = xlrd.open_workbook(input_path)
    #基本表
    basic_table = xls.sheet_by_name(sheet_basic_info)
    out_sheet_pend = basic_table.cell_value(1, 5)
    out_sheet_info = basic_table.cell_value(2, 5)
    out_sheet_chck = basic_table.cell_value(3, 5)
    out_sheet_chck_adu = basic_table.cell_value(4, 5)

    add_three_sheets(out_sheet_pend)
    add_three_sheets(out_sheet_info)
    add_three_sheets(out_sheet_chck)
    add_three_sheets(out_sheet_chck_adu)

def add_three_sheets(sheet_name):

    new_workbook = xls_file
    worksheet = new_workbook.add_sheet(sheet_name)
    xls = xlrd.open_workbook(input_path)

    hyperlink(worksheet)

    # 将"表字段"表 加入当前表
    length = 0  # 长度
    table1 = xls.sheet_by_name(sheet_field_str)
    length = write_to_worksheet(worksheet, table1, sheet_name, length)

    # 将"主键信息"表 加入当前表
    table2 = xls.sheet_by_name(prim_key_info)
    length = write_to_worksheet(worksheet, table2, sheet_name, length)

    # 将"表索引"表 加入当前表
    table3 = xls.sheet_by_name(sheet_idx_str)
    length = write_to_worksheet(worksheet, table3, sheet_name, length)

    # 将"分布键信息"表 加入当前表
    table4 = xls.sheet_by_name(distri_key_info)
    length = write_to_worksheet(worksheet, table4, sheet_name, length)

    # 将"接口对应文件"表 加入当前表
    table5 = xls.sheet_by_name(inf_cors_file)
    write_to_worksheet(worksheet, table5, sheet_name,length)
    new_workbook.save(output_path_final)

def write_to_worksheet(worksheet, table, sheet_name, length):

    head_style = pend_sheet_head_style()
    normol_style = pend_sheet_normal_style()

    rows = table.nrows
    cols = table.ncols
    for i in range(0, rows):
        table_name = table.cell_value(i, 1)
        if table_name == sheet_name or i == 0:
            length += 1
        for j in range(0, cols):
            if i == 0:
                worksheet.write(length, j, table.cell_value(i, j), head_style)
            if table_name == sheet_name:
                worksheet.write(length, j, table.cell_value(i, j), normol_style)
    length+=2
    for k in range(0, cols):
        worksheet.col(k).width = 180 * 43
        tall_style = xlwt.easyxf('font:height 256')  # 设置行高
        first_row = worksheet.row(k)
        first_row.set_style(tall_style)
    return length

# 添加超链接
def hyperlink(ws):
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.underline = 0x01
    font.name = "宋体"
    font.height = 10 * 20
    font.colour_index = 0x0C
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    style.font = font
    style.borders = borders
    link = 'HYPERLINK("#{}!A1";"{}")'.format(sheet_basic_info, back_str)
    ws.write(0, 0, xlwt.Formula(link), style)

# sheet3 head样式
def pend_sheet_head_style():
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = '宋体'
    font.height = 10 * 20
    font.bold = True
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = xlwt.Style.colour_map['pale_blue']
    style.pattern = pattern
    style.borders = borders
    style.font = font
    return style

def pend_sheet_normal_style():
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = '宋体'
    font.height = 10 * 20
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    style.borders = borders
    style.font = font
    return style

# 第五问
def add_exceplog_sheet():

    new_workbook = xls_file
    worksheet = new_workbook.add_sheet(exceplog_str)

    data = pd.ExcelFile(output_path_final)
    sheet_names = data.sheet_names
    ListZ = []
    for i in range(2, 6):
        sheetname = sheet_names[i]
        li = get_exceplog_list(sheetname)
        li2 = get_prim_index_exceplist(sheetname)
        ListZ.append(li)
        ListZ.append(li2)

    normol_style = pend_sheet_normal_style()
    yc = ['序号', '异常描述']
    worksheet.write(0, 0, yc[0], normol_style)
    worksheet.write(0, 1, yc[1], normol_style)

    idx = 0
    for li in ListZ:
        for str in li:
            idx += 1
            worksheet.write(idx, 0, idx, normol_style)
            worksheet.write(idx, 1, str, normol_style)

    new_workbook.save(output_path_final)

# 获取当前sheet异常
def get_exceplog_list(sheetname):

    df = pd.read_excel(output_path_final, sheet_name=sheetname)
    shape = df.shape
    # 找到key所在行和列
    keys = []
    tag = 0
    tag1 = 0
    for j in range(0, shape[1]):
        if not pd.isnull(df.iloc[0, j]):
            keys.append(df.iloc[0, j])
        else:
            break
    for i in range(0, shape[0]):
        if (pd.isnull(df.iloc[i, 0])):
            tag += 1
            if tag == 2:
                tag1 = i - 1
                break

    log_list = []
    for i in range(0, tag1):
        if (df.iloc[i, 10] == 'N-否'):
            for j in range(0, shape[1]):
                if (pd.isnull(df.iloc[i, j])):
                    logstr = '表"{}"中第{}行"{}"信息为空'.format(sheetname, i, keys[i])
                    log_list.append(logstr)
    return log_list

# 第六问
# 获取当前表主键异常和索引异常
def get_prim_index_exceplist(sheetname):

    df = pd.read_excel(output_path_final, sheet_name=sheetname)
    shape = df.shape
    keys = []
    rows_idx_list = []
    tag = 0
    for i in range(0, shape[0]):
        if (df.iloc[i, 1] == sheet_name_str):
            if i == 0: continue
            tag += 1
            if tag > 3 : break
            key = []
            for j in range(0, shape[1]):
                if not pd.isnull(df.iloc[i, j]):
                    key.append(df.iloc[i, j])
                else: break
            rows_idx_list.append(i)
            keys.append(key)
    # print(keys)
    # print(rows_idx_list)
    log_primkey_list = []
    for i in range(rows_idx_list[0], rows_idx_list[1]-2):
        if (pd.isnull(df.iloc[i, 2])):
            logstr = '{}表没有主键'.format(sheetname)
            log_primkey_list.append(logstr)
    for i in range(rows_idx_list[1], rows_idx_list[2]-2):
        if (pd.isnull(df.iloc[i, 2])):
            logstr = '{}表缺少索引'.format(sheetname)
            log_primkey_list.append(logstr)
    return log_primkey_list

if __name__ == '__main__':
    df1 = get_zero_data(input_path)
    write_sheet_zero(df1, output_path_final)
    copy_basic_sheet()
    add_all_three_sheets()
    add_exceplog_sheet()

