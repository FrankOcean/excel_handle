# -*- coding: utf-8 -*-
# @Author : PuYang
# @Time   : 2022/10/29
# @File   : opts.py
import os
import datetime
import pwd

input_path = 'resource/input_file.xls'
output_path = 'resource/output_file.xls'
output_path_final = 'output/output_file.xls'
cover_str = "封面"
sheet_field_str = "表字段"
prim_key_info = "主键信息"
sheet_idx_str = "表索引"
distri_key_info = "分布键信息"
inf_cors_file = "接口对应文件"
sheet_name_str = "表名称"
sheet_basic_info = "表基本信息"
back_str = "返回"

system_name_str = "系统名称"
subsys_tag_str = "子系统标识"
database_info_str = "数据库信息"

generate_user_str = "生成用户:"
generate_date_str = "生成日期:"
sheet_counts_str = "表数量:"

exceplog_str = "异常日志"

primkey_name = "主键名称"
index_name = "索引名称"

# 获取文件创建时间
def get_current_create_time():
    now_time = datetime.datetime.now()
    return "{}-{}-{}".format(now_time.year, now_time.month, now_time.day)

# 获取文件创建者
def get_owner(filename):
    stat = os.lstat(filename)
    uid = stat.st_uid
    pw = pwd.getpwuid(uid)
    return pw.pw_name

"""获取单元格背景颜色"""
def get_bgcolor(Book, sheet, row, col):
    xfx = sheet.cell_xf_index(row, col)
    xf = Book.xf_list[xfx]
    bgx = xf.background.pattern_colour_index
    return bgx

if __name__ == '__main__':
    import xlrd
    Excel = xlrd.open_workbook(input_path, formatting_info=True)
    sheet0 = Excel.sheets()[0]
    print(get_bgcolor(Excel, sheet0, 0, 0))