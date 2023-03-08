#!/usr/bin/python3.10
# -*- coding: utf-8 -*-
#
# Copyright (C) 2022 #
# @Time    : 2022/10/12 10:42
# @Author  : Sean
# @Email   : 756808193@qq.com

import pandas as pd
from utilFunc import *
from settings import filename1
from settings import filename2
from settings import filename3
from settings import filename4
from settings import filename5
from settings import filename6
from settings import filename7
from settings import filename8
from settings import filename9
from settings import filename10


#资产负债表配置
def merge_balance_sheet():

    df1=balance_sheet(filename1,'资产负债表')
    df2=balance_sheet(filename2,'资产负债表')
    df3=balance_sheet(filename3,'资产负债表',shou_tuo_zi_jin=83)
    df4=balance_sheet(filename4,'资产负债表',shou_tuo_zi_jin=83)
    df5=balance_sheet(filename5,'1001',shou_tuo_zi_jin=83)
    df6=balance_sheet(filename6,'资产负债表',shou_tuo_zi_jin=83)
    df7=balance_sheet(filename9,'1001',shou_tuo_zi_jin=83)

#统一列名
    df2.columns=df1.columns
    df3.columns=df1.columns
    df4.columns=df1.columns
    df5.columns=df1.columns
    df6.columns=df1.columns
    df7.columns=df1.columns

    # 连接多表
    df=pd.concat([df1,df2,df3,df4,df5,df6,df7],ignore_index=True)
    return df




#利润表配置
def merge_profit_sheet():
    #ben_yue_shu指本月数列的列数，ben_nian_lei_ji指本年累计数列的列数
    df1 = profit_sheet(filename1,'利润表')
    df2 = profit_sheet(filename2,'利润表',ben_yue_shu=1,ben_nian_lei_ji=2,hebing=True)
    df3 = profit_sheet(filename3, '利润表',bianzhi_row=1,project_row=2,start_row=3,end_row=35)
    df4 = profit_sheet(filename4, '利润表',bianzhi_row=1,project_row=2,start_row=3,end_row=35)
    df5 = profit_sheet(filename5, '1002',bianzhi_row=1,project_row=2,start_row=3,end_row=35)
    df6 = profit_sheet(filename6, '利润表')
    # df7 = profit_sheet(filename7, '1002',bianzhi_row=1,project_row=2,start_row=3,end_row=35)
    df8 = profit_sheet(filename8, '投行板块汇总-利润','A:E',start_row=4,end_row=29)
    df9 = profit_sheet(filename8, '投行板块华南区域总部-利润','A:E',start_row=4,end_row=29)
    df10 = profit_sheet(filename8, '投行板块上海区域总部-利润','A:E',start_row=4,end_row=29)
    df11 = profit_sheet(filename8, '投行板块北京区域总部-利润','A:E',start_row=4,end_row=29)
    df12 = profit_sheet(filename8, '债务融资总部-利润','A:E',start_row=4,end_row=29)
    df13 = profit_sheet(filename8, '中小企业融资部-利润','A:E',start_row=4,end_row=29)
    df14 = profit_sheet(filename8, '国际业务部-利润','A:E',start_row=4,end_row=29)
    df15 = profit_sheet(filename8, '资本市场部-利润','A:E',start_row=4,end_row=29)
    df16 = profit_sheet(filename8, '投行管理部门-利润','A:E',start_row=4,end_row=29)
    df17 = profit_sheet(filename8, '自营板块汇总-利润','A:E',start_row=4,end_row=29)
    df18 = profit_sheet(filename8, '自营-利润','A:E',start_row=4,end_row=29)
    df19 = profit_sheet(filename8, '衍生品交易部-利润','A:E',start_row=4,end_row=29)
    df20 = profit_sheet(filename8, '股转做市业务部-利润','A:E',start_row=4,end_row=29)
    df21 = profit_sheet(filename8, '财富板块-利润','A:E',start_row=4,end_row=29)
    df22 = profit_sheet(filename8, '受托客户资产管理业务-利润','A:E',start_row=4,end_row=29)
    df23 = profit_sheet(filename8, '质押融资部-利润','A:E',start_row=4,end_row=29)
    df24 = profit_sheet(filename8, '研发业务-利润','A:E',start_row=4,end_row=29)
    df25 = profit_sheet(filename8, '托管业务部-利润','A:E',start_row=4,end_row=29)
    df26 = profit_sheet(filename8, '总部-利润','A:E',start_row=4,end_row=29)
    df27 = profit_sheet(filename9, '1002',bianzhi_row=1,project_row=2,start_row=3,end_row=35)

    # 统一列名

    df2.columns = df1.columns
    df3.columns = df1.columns
    df4.columns = df1.columns
    df5.columns = df1.columns
    df6.columns = df1.columns
    # df7.columns = df1.columns
    # df8.columns = df1.columns
    # df9.columns = df1.columns
    # df10.columns = df1.columns
    # df11.columns = df1.columns
    # df12.columns = df1.columns
    # df13.columns = df1.columns
    # df14.columns = df1.columns
    # df15.columns = df1.columns
    # df16.columns = df1.columns
    # df17.columns = df1.columns
    # df18.columns = df1.columns
    # df19.columns = df1.columns
    # df20.columns = df1.columns
    # df21.columns = df1.columns
    # df22.columns = df1.columns
    # df23.columns = df1.columns
    # df24.columns = df1.columns
    # df25.columns = df1.columns
    # df26.columns = df1.columns
    df27.columns = df1.columns


    df = pd.concat([df1, df2, df3, df4, df5, df6,df8,df9,df10,df11,df12,df13,df14,df15,df16,df17,df18,df19,df20,df21,df22,df23,df24,df25,df26,df27], ignore_index=True)
    return df

#业务管理费用表配置
def merge_bussiness_management_fee_sheet():
    df1=bussiness_management_fee_sheet(filename1,'业务及管理费用表')
    df2=bussiness_management_fee_sheet(filename2,'费用表')
    df3=bussiness_management_fee_sheet(filename3,'业务及管理费用表')
    df4=bussiness_management_fee_sheet(filename4,'业务及管理费明细表')
    df5=bussiness_management_fee_sheet(filename5,'1003',bianzhi_row=1,project_row=2,start_row=3,end_row=48)
    df6=bussiness_management_fee_sheet(filename6,'业务及管理费用表')
    # df7=bussiness_management_fee_sheet(filename7,'1003',bianzhi_row=1)
    df8 = bussiness_management_fee_sheet(filename8, '投行板块汇总-费用',usecols='A:E')
    df9 = bussiness_management_fee_sheet(filename8, '投行板块华南区域总部-费用',usecols='A:E')
    df10 = bussiness_management_fee_sheet(filename8, '投行板块上海区域总部-费用',usecols='A:E')
    df11 = bussiness_management_fee_sheet(filename8, '投行板块北京区域总部-费用',usecols='A:E')
    df12 = bussiness_management_fee_sheet(filename8, '债务融资总部-费用',usecols='A:E')
    df13 = bussiness_management_fee_sheet(filename8, '中小企业融资部-费用',usecols='A:E')
    df14 = bussiness_management_fee_sheet(filename8, '国际业务部-费用',usecols='A:E')
    df15 = bussiness_management_fee_sheet(filename8, '资本市场部-费用',usecols='A:E')
    df16 = bussiness_management_fee_sheet(filename8, '投行管理部门-费用',usecols='A:E')
    df17 = bussiness_management_fee_sheet(filename8, '自营板块汇总-费用',usecols='A:E')
    df18 = bussiness_management_fee_sheet(filename8, '自营-费用',usecols='A:E')
    df19 = bussiness_management_fee_sheet(filename8, '衍生品交易部-费用',usecols='A:E')
    df20 = bussiness_management_fee_sheet(filename8, '股转做市业务部-费用',usecols='A:E')
    df21 = bussiness_management_fee_sheet(filename8, '财富板块-费用',usecols='A:E')
    df22 = bussiness_management_fee_sheet(filename8, '受托客户资产管理业务-费用',usecols='A:E')
    df23 = bussiness_management_fee_sheet(filename8, '质押融资部-费用',usecols='A:E')
    df24 = bussiness_management_fee_sheet(filename8, '研发业务-费用',usecols='A:E')
    df25 = bussiness_management_fee_sheet(filename8, '托管业务部-费用',usecols='A:E')
    df26 = bussiness_management_fee_sheet(filename8, '总部-费用',usecols='A:E')
    df27 = bussiness_management_fee_sheet(filename9, '1003',bianzhi_row=1)

    # 统一列名

    df2.columns = df1.columns
    df3.columns = df1.columns
    df4.columns = df1.columns
    df5.columns = df1.columns
    df6.columns = df1.columns
    # df7.columns = df1.columns
    df8.columns = df1.columns
    df9.columns = df1.columns
    df10.columns = df1.columns
    df11.columns = df1.columns
    df12.columns = df1.columns
    df13.columns = df1.columns
    df14.columns = df1.columns
    df15.columns = df1.columns
    df16.columns = df1.columns
    df17.columns = df1.columns
    df18.columns = df1.columns
    df19.columns = df1.columns
    df20.columns = df1.columns
    df21.columns = df1.columns
    df22.columns = df1.columns
    df23.columns = df1.columns
    df24.columns = df1.columns
    df25.columns = df1.columns
    df26.columns = df1.columns
    df27.columns = df1.columns
    df = pd.concat(
        [df1, df2, df3, df4, df5, df6, df8, df9, df10, df11, df12, df13, df14, df15, df16, df17, df18, df19, df20,
         df21, df22, df23, df24, df25, df26, df27], ignore_index=True)
    return df

#资产负债表附表配置
def merge_balance_abundant_sheet():
    # df1 = balance_abundant_sheet(filename7, '2001')
    df2 = balance_abundant_sheet(filename9, '2001')

    # 统一列名
    # df2.columns = df1.columns
    df = pd.concat([df2],ignore_index=True)
    for i in range(len(df['编制单位'])) :
        if df.loc[i, '编制单位'] == None or df.loc[i, '编制单位'] == "":
            continue
        if df.loc[i,'编制单位'].find(':') != -1:
            df.loc[i,'编制单位']=df.loc[i,'编制单位'].split(':')[1]
        elif df.loc[i,'编制单位'].find('：') != -1:
            df.loc[i, '编制单位'] = df.loc[i, '编制单位'].split('：')[1]
        else:
            continue
    return df

#利润表附表配置
def merge_profit_abundant_sheet():
    # df1 = profit_abundant_sheet(filename7, '2002')

    # 统一列名
    # df2.columns = df1.columns

    df = pd.concat([df2], ignore_index=True)

    for i in range(len(df['编制单位'])):
        if df.loc[i, '编制单位'] == None or df.loc[i, '编制单位'] == "":
            continue
        if df.loc[i, '编制单位'].find(':') != -1:
            df.loc[i, '编制单位'] = df.loc[i, '编制单位'].split(':')[1]
        elif df.loc[i, '编制单位'].find('：') != -1:
            df.loc[i, '编制单位'] = df.loc[i, '编制单位'].split('：')[1]
        else:
            continue

    return df

#业务管理费用附表配置
def merge_bussiness_management_fee_abundant_sheet():
    df1 = bussiness_management_fee_abundant_sheet(filename3, '业务及管理费附表')
    #没有编制单位从前一个取
    df_bian_zhi_1 = profit_sheet(filename3, '利润表', bianzhi_row=1, project_row=2, start_row=3, end_row=35)
    df1['编制单位']=df_bian_zhi_1['编制单位']
    df2 = bussiness_management_fee_abundant_sheet(filename4, '业务及管理费附表')
    # 没有编制单位从前一个取
    df_bian_zhi_2 = profit_sheet(filename4, '利润表', bianzhi_row=1, project_row=2, start_row=3, end_row=35)
    df2['编制单位'] = df_bian_zhi_2['编制单位']
    df3 = bussiness_management_fee_abundant_sheet(filename5, '管理费用表附表')
    # 没有编制单位从前一个取
    df_bian_zhi_3 = profit_sheet(filename5, '1002', bianzhi_row=1, project_row=2, start_row=3, end_row=35)
    df3['编制单位'] = df_bian_zhi_3['编制单位']
    df4 = bussiness_management_fee_abundant_sheet(filename6, '业务及管理费附表')
    # 没有编制单位从前一个取
    df_bian_zhi_4 = profit_sheet(filename6, '利润表')
    df4['编制单位'] = df_bian_zhi_4['编制单位']

    df5 = bussiness_management_fee_abundant_sheet(filename9, '2003')


    #统一列名
    df2.columns = df1.columns
    df3.columns = df1.columns
    df4.columns = df1.columns
    df5.columns = df1.columns

    #连接多表
    df = pd.concat([df1, df2, df3, df4, df5], ignore_index=True)

    for i in range(len(df['编制单位'])):
        if type(df.loc[i, '编制单位']) != type(""):

            df.loc[i, '编制单位']=""
            continue
        if df.loc[i, '编制单位'].find(':') != -1:
            df.loc[i, '编制单位'] = df.loc[i, '编制单位'].split(':')[1]
        elif df.loc[i, '编制单位'].find('：') != -1:
            df.loc[i, '编制单位'] = df.loc[i, '编制单位'].split('：')[1]
        else:
            continue

    return df

#净资本计算表配置
def merge_capital_caculation_sheet():
    df1 = capital_caculation_sheet(filename10, '净资本计算表')

    # for i in range(len(df1['编制单位'])):
    #     if df1.loc[i, '编制单位'].find(':') != -1:
    #         df1.loc[i, '编制单位'] = df1.loc[i, '编制单位'].split(':')[1]
    #     elif df1.loc[i, '编制单位'].find('：') != -1:
    #         df1.loc[i, '编制单位'] = df1.loc[i, '编制单位'].split('：')[1]
    #     else:
    #         continue
    return df1

df1=merge_balance_sheet()
df2=merge_profit_sheet()
df3=merge_bussiness_management_fee_sheet()
df4=merge_balance_abundant_sheet()
df5=merge_profit_abundant_sheet()
df6=merge_bussiness_management_fee_abundant_sheet()
df7=merge_capital_caculation_sheet()

with pd.ExcelWriter('test.xlsx') as writer:
    df1.to_excel(writer, sheet_name='资产负债表（1001）', index=False)
    df2.to_excel(writer, sheet_name='利润表（1002）', index=False)
    df3.to_excel(writer, sheet_name='业务及管理费用表（1003）', index=False)
    df4.to_excel(writer, sheet_name='资产负债表附表（2001）', index=False)
    df5.to_excel(writer, sheet_name='利润表附表（2002）', index=False)
    df6.to_excel(writer, sheet_name='业务及管理费用附表（2003）', index=False)
    df7.to_excel(writer, sheet_name='证券公司净资本计算表', index=False)