#!/usr/bin/python3.10
# -*- coding: utf-8 -*-
#
# Copyright (C) 2022 #
# @Time    : 2022/10/12 10:42
# @Author  : Sean
# @Email   : 756808193@qq.com
import pandas as pd
from settings import month


#显示所有列
pd.set_option('display.max_columns', None)
#显示所有行
pd.set_option('display.max_rows', None)

#一般列到列，列数，编制单位，项目位置，数据位置都是不变的，这些可以写死在对应的函数里，变的是excel_file和sheet名

#资产负债表处理函数
def balance_sheet(excel_file, sheet_name, start_row_1=5,end_row_1=45,start_row_2=46,end_row_2=81,start_row_3=5,end_row_3=31,start_row_4=44,end_row_4=45,
start_row_5=46,end_row_5=54,start_row_6=60,end_row_6=61,start_row_7=68,end_row_7=81,huo_bi_zi_jin=82,shou_tuo_tou_zi=83,shou_tuo_zi_jin=82,project_row=3,bianzhi_row=2):

    df=pd.read_excel(excel_file,sheet_name=sheet_name,usecols='B:I',header=None,names=range(8))

    #固定项目声明
    column_list = ['月份', '编制单位', '项目']
    global month

    append_column = df.iloc[start_row_1:end_row_1, 0].values.tolist() + df.iloc[start_row_2:end_row_2, 0].values.tolist() + df.iloc[start_row_3:end_row_3, 0].values.tolist()
    append_column.append(df.iloc[start_row_4, 4])
    append_column = append_column + df.iloc[start_row_5:end_row_5, 4].values.tolist()
    append_column.append(df.iloc[start_row_6, 4])
    append_column = append_column + df.iloc[start_row_7:end_row_7, 4].values.tolist()
    append_column.append(df.iloc[huo_bi_zi_jin, 0])
    append_column.append(df.iloc[shou_tuo_tou_zi, 0])
    append_column.append(df.iloc[shou_tuo_zi_jin, 4])
    # 去空格

    append_column = [i.strip() for i in append_column]
    column_list = column_list + append_column


    #年初数
    value_list = []
    #月份
    value_list.append(month)
    #编制单位名和项目名的值
    if df.iloc[2, 0].find(':') !=-1 :
        value_list.extend([df.iloc[2, 0].split(':')[1], df.iloc[3, 2]])
    elif df.iloc[2, 0].find('：') !=-1:
        value_list.extend([df.iloc[2, 0].split('：')[1], df.iloc[3, 2]])
    else:
        value_list.extend([df.iloc[2, 0], df.iloc[3, 2]])
    #项目具体值
    # value_list = value_list + df.iloc[5:81, 2].values.tolist() + df.iloc[5:31, 6].values.tolist()
    # value_list.append(df.iloc[44, 6])
    # value_list = value_list + df.iloc[46:54, 6].values.tolist()
    # value_list.append(df.iloc[60, 6])
    # value_list = value_list + df.iloc[68:81, 6].values.tolist() + df.iloc[82:84, 2].values.tolist()
    # value_list.append(df.iloc[82, 6])

    value_list = value_list + df.iloc[start_row_1:end_row_1, 2].values.tolist() + df.iloc[start_row_2:end_row_2,
                                                                        2].values.tolist() + df.iloc[
                                                                                             start_row_3:end_row_3,
                                                                                             6].values.tolist()
    value_list.append(df.iloc[start_row_4, 6])
    value_list = value_list + df.iloc[start_row_5:end_row_5, 6].values.tolist()
    value_list.append(df.iloc[start_row_6, 6])
    value_list = value_list + df.iloc[start_row_7:end_row_7, 6].values.tolist()
    value_list.append(df.iloc[huo_bi_zi_jin, 2])
    value_list.append(df.iloc[shou_tuo_tou_zi, 2])
    value_list.append(df.iloc[shou_tuo_zi_jin, 6])


    #期末数
    value_list_2 = []
    # 月份
    value_list_2.append(month)
    # 编制单位名和项目名的值
    if df.iloc[2, 0].find(':') != -1:
        value_list_2.extend([df.iloc[2, 0].split(':')[1], df.iloc[3, 3]])
    elif df.iloc[2, 0].find('：') != -1:
        value_list_2.extend([df.iloc[2, 0].split('：')[1], df.iloc[3, 3]])
    else:
        value_list_2.extend([df.iloc[2, 0], df.iloc[3, 2]])

    # 项目具体值
    # value_list_2 = value_list_2 + df.iloc[5:81, 3].values.tolist() + df.iloc[5:31, 7].values.tolist()
    # value_list_2.append(df.iloc[44, 7])
    # value_list_2 = value_list_2 + df.iloc[46:54, 7].values.tolist()
    # value_list_2.append(df.iloc[60, 7])
    # value_list_2 = value_list_2 + df.iloc[68:81, 7].values.tolist() + df.iloc[82:84, 3].values.tolist()
    # value_list_2.append(df.iloc[82, 7])

    value_list_2 = value_list_2+ df.iloc[start_row_1:end_row_1, 3].values.tolist() + df.iloc[start_row_2:end_row_2,
                                                                     3].values.tolist() + df.iloc[
                                                                                          start_row_3:end_row_3,
                                                                                          3].values.tolist()
    value_list_2.append(df.iloc[start_row_4, 7])
    value_list_2 = value_list_2 + df.iloc[start_row_5:end_row_5, 7].values.tolist()
    value_list_2.append(df.iloc[start_row_6, 7])
    value_list_2 = value_list_2 + df.iloc[start_row_7:end_row_7, 7].values.tolist()
    value_list_2.append(df.iloc[huo_bi_zi_jin, 3])
    value_list_2.append(df.iloc[shou_tuo_tou_zi, 3])
    value_list_2.append(df.iloc[shou_tuo_zi_jin, 7])


    df_all = pd.DataFrame([value_list, value_list_2], columns=column_list)
    df_all = df_all.fillna(0)
    df_all
    return df_all


#利润表表处理函数
def profit_sheet(excel_file, sheet_name,usecols='B:F',start_row=4,end_row=36,project_row=3,bianzhi_row=2,ben_yue_shu=2,ben_nian_lei_ji=3,ben_nian_ji_hua=1,wan_cheng_ji_hua=4,hebing=False):

    df = pd.read_excel(excel_file,sheet_name=sheet_name,usecols=usecols,header=None,names=range(5))

    column_list=['月份','编制单位','项目']

    global month

    append_column=df.iloc[start_row:end_row,0].values.tolist()
    #去空格
    append_column =[i.strip() for i in append_column]
    column_list=column_list+append_column


    #本月数
    value_list=[]
    value_list.append(month)
    # 编制单位名和项目名的值
    if df.iloc[bianzhi_row, 0].find(':') != -1:
        value_list.extend([df.iloc[bianzhi_row, 0].split(':')[1], df.iloc[project_row,ben_yue_shu]])
    elif df.iloc[bianzhi_row, 0].find('：') != -1:
        value_list.extend([df.iloc[bianzhi_row, 0].split('：')[1], df.iloc[project_row, ben_yue_shu]])
    else:
        value_list.extend([df.iloc[bianzhi_row, 0], df.iloc[project_row, 2]])


    value_list=value_list+df.iloc[start_row:end_row,ben_yue_shu].values.tolist()



    # 本年累计数
    value_list_2=[]
    value_list_2.append(month)
    if df.iloc[bianzhi_row, 0].find(':') != -1:
        value_list_2.extend([df.iloc[bianzhi_row, 0].split(':')[1], df.iloc[project_row, ben_nian_lei_ji]])
    elif df.iloc[bianzhi_row, 0].find('：') != -1:
        value_list_2.extend([df.iloc[bianzhi_row, 0].split('：')[1], df.iloc[project_row, ben_nian_lei_ji]])
    else:
        value_list_2.extend([df.iloc[bianzhi_row, 0], df.iloc[project_row, 3]])

    value_list_2=value_list_2+df.iloc[start_row:end_row,ben_nian_lei_ji].values.tolist()



    if hebing == False:
        #本年计划数

        value_list_3=[]
        value_list_3.append(month)
        if df.iloc[bianzhi_row, 0].find(':') != -1:
            value_list_3.extend([df.iloc[bianzhi_row, 0].split(':')[1], df.iloc[project_row, 1]])
        elif df.iloc[bianzhi_row, 0].find('：') != -1:
            value_list_3.extend([df.iloc[bianzhi_row, 0].split('：')[1], df.iloc[project_row, 1]])
        else:
            value_list_3.extend([df.iloc[bianzhi_row, 0], df.iloc[project_row, 1]])

        value_list_3=value_list_3+df.iloc[start_row:end_row,ben_nian_ji_hua].values.tolist()


        #完成计划
        value_list_4=[]
        value_list_4.append(month)
        if df.iloc[bianzhi_row, 0].find(':') != -1:
            value_list_4.extend([df.iloc[bianzhi_row, 0].split(':')[1], df.iloc[project_row, 4]])
        elif df.iloc[bianzhi_row, 0].find('：') != -1:
            value_list_4.extend([df.iloc[bianzhi_row, 0].split('：')[1], df.iloc[project_row, 4]])
        else:
            value_list_4.extend([df.iloc[bianzhi_row, 0], df.iloc[project_row, 4]])

        value_list_4=value_list_4+df.iloc[start_row:end_row,wan_cheng_ji_hua].values.tolist()


        df_all=pd.DataFrame([value_list,value_list_2,value_list_3,value_list_4],columns=column_list)
        df_all=df_all.fillna(0)
        df_all
        return df_all
    else:
        df_all = pd.DataFrame([value_list, value_list_2], columns=column_list)
        return df_all

#业务管理费用表处理函数
def bussiness_management_fee_sheet(excel_file, sheet_name,usecols='B:F',start_row=4,end_row=49,project_row=3,bianzhi_row=2,ben_yue_shi_ji=2,ben_nian_lei_ji=3,ben_nian_ji_hua=1,wan_cheng_ji_hua=4):
    df = pd.read_excel(excel_file, sheet_name=sheet_name, usecols=usecols, header=None, names=range(5))

    column_list = ['月份', '编制单位', '项目']
    append_column = df.iloc[start_row:end_row, 0].values.tolist()
    # 去空格
    append_column = [i.strip() for i in append_column]
    column_list = column_list + append_column
    len(column_list)
    global month
    #本月实际
    value_list = []
    value_list.append(month)
    if df.iloc[bianzhi_row, 0].find(':') != -1:
        value_list.extend([df.iloc[bianzhi_row, 0].split(':')[1], df.iloc[project_row, ben_yue_shi_ji]])
    elif df.iloc[bianzhi_row, 0].find('：') != -1:
        value_list.extend([df.iloc[bianzhi_row, 0].split('：')[1], df.iloc[project_row, ben_yue_shi_ji]])
    else:
        value_list.extend([df.iloc[bianzhi_row, 0], df.iloc[project_row, ben_yue_shi_ji]])

    value_list = value_list + df.iloc[start_row:end_row, ben_yue_shi_ji].values.tolist()
    len(value_list)
    value_list

    # 本年累计
    value_list_2 = []
    value_list_2.append(month)
    if df.iloc[bianzhi_row, 0].find(':') != -1:
        value_list_2.extend([df.iloc[bianzhi_row, 0].split(':')[1], df.iloc[project_row, ben_nian_lei_ji]])
    elif df.iloc[bianzhi_row, 0].find('：') != -1:
        value_list_2.extend([df.iloc[bianzhi_row, 0].split('：')[1], df.iloc[project_row, ben_nian_lei_ji]])
    else:
        value_list_2.extend([df.iloc[bianzhi_row, 0], df.iloc[project_row, ben_nian_lei_ji]])

    value_list_2 = value_list_2 + df.iloc[start_row:end_row, ben_nian_lei_ji].values.tolist()
    len(value_list_2)
    value_list_2

    # 本年计划
    value_list_3 = []
    value_list_3.append(month)
    if df.iloc[bianzhi_row, 0].find(':') != -1:
        value_list_3.extend([df.iloc[bianzhi_row, 0].split(':')[1], df.iloc[project_row, ben_nian_ji_hua]])
    elif df.iloc[bianzhi_row, 0].find('：') != -1:
        value_list_3.extend([df.iloc[bianzhi_row, 0].split('：')[1], df.iloc[project_row, ben_nian_ji_hua]])
    else:
        value_list_3.extend([df.iloc[bianzhi_row, 0], df.iloc[project_row, ben_nian_ji_hua]])
    value_list_3 = value_list_3 + df.iloc[start_row:end_row, ben_nian_ji_hua].values.tolist()
    len(value_list_3)
    value_list_3

    #完成计划
    value_list_4 = []
    value_list_4.append(month)
    if df.iloc[bianzhi_row, 0].find(':') != -1:
        value_list_4.extend([df.iloc[bianzhi_row, 0].split(':')[1], df.iloc[project_row, wan_cheng_ji_hua]])
    elif df.iloc[bianzhi_row, 0].find('：') != -1:
        value_list_4.extend([df.iloc[bianzhi_row, 0].split('：')[1], df.iloc[project_row, wan_cheng_ji_hua]])
    else:
        value_list_4.extend([df.iloc[bianzhi_row, 0], df.iloc[project_row, wan_cheng_ji_hua]])
    value_list_4 = value_list_4 + df.iloc[start_row:end_row, wan_cheng_ji_hua].values.tolist()
    len(value_list_4)
    value_list_4

    df_all = pd.DataFrame([value_list, value_list_2, value_list_3, value_list_4], columns=column_list)
    df_all = df_all.fillna(0)
    df_all
    return df_all

#资产负债表附表处理函数
def balance_abundant_sheet(excel_file, sheet_name):
    df = pd.read_excel(excel_file, sheet_name=sheet_name, usecols='A:E', header=None, names=range(5))


    column_list = ['月份', '编制单位', '项目']
    append_column = df.iloc[4:28, 0].values.tolist() + df.iloc[29:54, 0].values.tolist() + \
                    df.iloc[55:88,0].values.tolist() + df.iloc[89:99,0].values.tolist() + \
                    df.iloc[100:133,0].values.tolist() + df.iloc[134:184,0].values.tolist() + \
                    df.iloc[185:191,0].values.tolist() + df.iloc[192:206,0].values.tolist() + \
                    df.iloc[207:215,0].values.tolist()

    # 去空格
    append_column = [i.strip() for i in append_column]
    column_list = column_list + append_column
    len(column_list)

    global month

    value_list = []
    value_list.append(month)
    value_list.extend([df.iloc[2, 0], df.iloc[3, 1]])
    value_list = value_list + df.iloc[4:28, 1].values.tolist() + df.iloc[29:54, 1].values.tolist() + df.iloc[55:88,
                                                                                                     1].values.tolist() + df.iloc[
                                                                                                                          89:99,
                                                                                                                          1].values.tolist() + df.iloc[
                                                                                                                                               100:133,
                                                                                                                                               1].values.tolist() + df.iloc[
                                                                                                                                                                    134:184,
                                                                                                                                                                    1].values.tolist() + df.iloc[
                                                                                                                                                                                         185:191,
                                                                                                                                                                                         1].values.tolist() + df.iloc[
                                                                                                                                                                                                              192:206,
                                                                                                                                                                                                              1].values.tolist() + df.iloc[
                                                                                                                                                                                                                                   207:215,
                                                                                                                                                                                                                                   1].values.tolist()
    len(value_list)
    value_list

    value_list_2 = []
    value_list_2.append(month)
    value_list_2.extend([df.iloc[2, 0], df.iloc[3, 2]])
    value_list_2 = value_list_2 + df.iloc[4:28, 2].values.tolist() + df.iloc[29:54, 2].values.tolist() + df.iloc[55:88,
                                                                                                         2].values.tolist() + df.iloc[
                                                                                                                              89:99,
                                                                                                                              2].values.tolist() + df.iloc[
                                                                                                                                                   100:133,
                                                                                                                                                   2].values.tolist() + df.iloc[
                                                                                                                                                                        134:184,
                                                                                                                                                                        2].values.tolist() + df.iloc[
                                                                                                                                                                                             185:191,
                                                                                                                                                                                             2].values.tolist() + df.iloc[
                                                                                                                                                                                                                  192:206,
                                                                                                                                                                                                                  2].values.tolist() + df.iloc[
                                                                                                                                                                                                                                       207:215,
                                                                                                                                                                                                                                       2].values.tolist()
    len(value_list_2)
    value_list

    value_list_3 = []
    value_list_3.append(month)
    value_list_3.extend([df.iloc[2, 0], df.iloc[3, 3]])
    value_list_3 = value_list_3 + df.iloc[4:28, 3].values.tolist() + df.iloc[29:54, 3].values.tolist() + df.iloc[55:88,
                                                                                                         1].values.tolist() + df.iloc[
                                                                                                                              89:99,
                                                                                                                              3].values.tolist() + df.iloc[
                                                                                                                                                   100:133,
                                                                                                                                                   3].values.tolist() + df.iloc[
                                                                                                                                                                        134:184,
                                                                                                                                                                        3].values.tolist() + df.iloc[
                                                                                                                                                                                             185:191,
                                                                                                                                                                                             3].values.tolist() + df.iloc[
                                                                                                                                                                                                                  192:206,
                                                                                                                                                                                                                  3].values.tolist() + df.iloc[
                                                                                                                                                                                                                                       207:215,
                                                                                                                                                                                                                                       3].values.tolist()
    len(value_list_3)
    value_list

    value_list_4 = []
    value_list_4.append(month)
    value_list_4.extend([df.iloc[2, 0], df.iloc[3, 4]])
    value_list_4 = value_list_4 + df.iloc[4:28, 4].values.tolist() + df.iloc[29:54, 4].values.tolist() + df.iloc[55:88,
                                                                                                         4].values.tolist() + df.iloc[
                                                                                                                              89:99,
                                                                                                                              4].values.tolist() + df.iloc[
                                                                                                                                                   100:133,
                                                                                                                                                   4].values.tolist() + df.iloc[
                                                                                                                                                                        134:184,
                                                                                                                                                                        4].values.tolist() + df.iloc[
                                                                                                                                                                                             185:191,
                                                                                                                                                                                             4].values.tolist() + df.iloc[
                                                                                                                                                                                                                  192:206,
                                                                                                                                                                                                                  4].values.tolist() + df.iloc[
                                                                                                                                                                                                                                       207:215,
                                                                                                                                                                                                                                       4].values.tolist()
    len(value_list_4)
    value_list

    df_all = pd.DataFrame([value_list, value_list_2, value_list_3, value_list_4], columns=column_list)
    df_all = df_all.fillna(0)
    df_all
    return df_all

#利润表附表处理函数
def profit_abundant_sheet(excel_file, sheet_name):
    df = pd.read_excel(excel_file, sheet_name=sheet_name, usecols='B:E', header=None, names=range(5))

    column_list = ['月份', '编制单位', '项目']
    append_column = df.iloc[5:185, 0].values.tolist()

    # 去空格
    append_column = [i.strip() for i in append_column]
    column_list = column_list + append_column

    global month

    value_list = []
    value_list.append(month)
    value_list.extend([df.iloc[2, 0], df.iloc[3, 2]])
    value_list = value_list + df.iloc[5:185, 2].values.tolist()


    value_list_2 = []
    value_list_2.append(month)
    value_list_2.extend([df.iloc[2, 0], df.iloc[3, 3]])
    value_list_2 = value_list_2 + df.iloc[5:185, 3].values.tolist()


    value_list_3 = []
    value_list_3.append(month)
    value_list_3.extend([df.iloc[2, 0], df.iloc[3, 1]])
    value_list_3 = value_list_3 + df.iloc[5:185, 1].values.tolist()


    #     value_list_4=[]
    #     value_list_4.append(month)
    #     value_list_4.extend([df.iloc[2,0],df.iloc[3,4]])
    #     value_list_4=value_list_4+df.iloc[4:28,4].values.tolist()+df.iloc[29:54,4].values.tolist()+df.iloc[55:88,4].values.tolist()+df.iloc[89:99,4].values.tolist()+df.iloc[100:133,4].values.tolist()+df.iloc[134:184,4].values.tolist()+df.iloc[185:191,4].values.tolist()+df.iloc[192:206,4].values.tolist()+df.iloc[207:215,4].values.tolist()
    #     len(value_list_4)
    #     value_list

    df_all = pd.DataFrame([value_list, value_list_2, value_list_3], columns=column_list)
    df_all = df_all.fillna(0)
    df_all
    return df_all

#业务管理费用附表处理函数
def bussiness_management_fee_abundant_sheet(excel_file, sheet_name):
    df = pd.read_excel(excel_file, sheet_name=sheet_name, usecols='A:D', header=None, names=range(4))


    column_list = ['月份', '编制单位', '项目']
    append_column = df.iloc[4:95, 1].values.tolist()

    # 去空格
    append_column = [i.strip() for i in append_column]
    column_list = column_list + append_column
    len(column_list)

    value_list = []
    global month
    value_list.append(month)
    value_list.extend([df.iloc[2, 1], df.iloc[3, 2]])
    value_list = value_list + df.iloc[4:95, 2].values.tolist()
    len(value_list)
    value_list

    value_list_2 = []
    value_list_2.append(month)
    value_list_2.extend([df.iloc[2, 1], df.iloc[3, 3]])
    value_list_2 = value_list_2 + df.iloc[4:95, 3].values.tolist()
    len(value_list_2)
    value_list_2

    #     value_list_3=[]
    #     value_list_3.append(month)
    #     value_list_3.extend([df.iloc[2,0],df.iloc[3,1]])
    #     value_list_3=value_list_3+df.iloc[4:95,1].fillna(0).values.tolist()
    #     len(value_list_3)
    #     value_list_3

    #     value_list_4=[]
    #     value_list_4.append(month)
    #     value_list_4.extend([df.iloc[2,0],df.iloc[3,4]])
    #     value_list_4=value_list_4+df.iloc[4:95,4].fillna(0).values.tolist()
    #     len(value_list_4)
    #     value_list_4

    df_all = pd.DataFrame([value_list, value_list_2], columns=column_list)
    df_all = df_all.fillna(0)
    df_all
    return df_all

#净资本计算表处理函数
def capital_caculation_sheet(excel_file, sheet_name):
    df = pd.read_excel(excel_file, sheet_name=sheet_name, usecols='A:E', header=None, names=range(5))

    column_list = ['月份', '项目']
    append_column = df.iloc[4:28, 0].values.tolist()


    # 去空格
    append_column = [i.strip() for i in append_column]
    column_list = column_list + append_column
    len(column_list)

    value_list = []
    global month
    value_list.append(month)
    value_list.append(df.iloc[2, 2])
    value_list = value_list + df.iloc[4:28, 2].values.tolist()

    value_list_2 = []
    value_list_2.append(month)
    value_list_2.append( df.iloc[2, 3])
    value_list_2 = value_list_2 + df.iloc[4:28, 3].values.tolist()

    value_list_3 = []
    value_list_3.append(month)
    value_list_3.append(df.iloc[2, 4])
    value_list_3 = value_list_3 + df.iloc[4:28, 4].values.tolist()


    df_all = pd.DataFrame([value_list, value_list_2,value_list_3], columns=column_list)
    df_all = df_all.fillna(0)
    df_all
    return df_all


