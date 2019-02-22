# 肿物大小
# -*- coding:utf-8 -*-
import xlrd
import numpy
import re
from xlutils.copy import copy
from functools import singledispatch


CNN=['一','二','两','三','四','五','六','七','八','九']
numb=[1,2,2,3,4,5,6,7,8,9]

def changeNumb(string):
    tmp=string
    for i in range(len(CNN)):
        if CNN[i] in tmp:
            tmp=tmp.replace(CNN[i],str(numb[i]))
    return tmp

@singledispatch
def size(key):
    ans=[]
    string=''
    tmp_key = key.replace('(', '（')
    tmp_key = key.replace('×', '*')
    tmp_key = tmp_key.replace(')', '）')
    p3 = r'[^。!！;；]+'
    pattern3=re.compile(p3)
    sentence=pattern3.findall(tmp_key)
    for tmp_key2 in sentence:
        if '未' in tmp_key2:
            continue
        if '灶' in tmp_key2 or '肿物' in tmp_key2 or '肿块' in tmp_key2 or '结节' in tmp_key2:
            p1 = r'[^,，。:：“”‘’!！;；]+'
            pattren1 = re.compile(p1)
            a = pattren1.findall(tmp_key2)
            p = r'[\d\.]+[Cc][Mm]\*[\d\.]+[Cc][Mm]\*[\d\.]+[Cc][Mm]|[\d\.]+[Cc][Mm]\*[\d\.]+[Cc][Mm]|[\d\.]+[Cc][Mm]至[\d\.]+[Cc][Mm]|不足[\d\.]+[Cc][Mm]|[\d\.]+[Cc][Mm]'
            pattern = re.compile(p)
            tumor = pattern.findall(tmp_key2)
            for target in tumor:
                for i in range(len(a)):
                    where = ''
                    if target in a[i]:
                        if '见' in a[i]:
                            where = a[i].split('见')[0]
                            if where != '':
                                where += '见'
                        elif i > 0 and '见' in a[i - 1]:
                            where = a[i - 1].split('见')[0]
                            if where != '':
                                where += '见'
                        if target in where:
                            continue
                        if where != '':
                            if '其' in where or '切面' in where:
                                ans.append(target)
                            else:
                                ans.append(target + '(' + where + ')')
                        else:
                            ans.append(target)
    if ans!=[]:
        for i in ans:
            string+=i+'、'
    string=string[:len(string)-1]
    return string

@size.register(list)
def _size(sentence):
    ans=[]
    string=''
    for key in sentence:
        tmp=size(key)
        if tmp!='' and tmp!=[]:
            ans.append(tmp)
    if ans!=[]:
        for i in ans:
            string+=i+'、'
    string=string[:len(string)-1]
    return string

# key='1.结合临床,直肠癌术后,双肺转移瘤,较前2016-6-9变化不著 2.双侧胸膜略增厚,变化不著 4.甲状腺低密度灶,变化不著。双肺近胸膜下示多个结节灶及肿块,部分边缘分叶、毛糙,牵拉邻近胸膜,大者位于左肺下叶,长径约3.0CM(纵隔窗),内示小空泡影,增强呈不均质强化。双侧肺门及纵隔未见增大淋巴结。双侧胸膜略增厚。双侧胸腔未见积液征象。甲状腺右叶见低密度灶,边缘较清晰。	'
# print(size(key))


Query_excel = xlrd.open_workbook('原始数据文本.xlsx')
Save_excel = copy(Query_excel)
# 获得sheet的对象
Query_excel_sheet = Query_excel.sheet_by_name('Sheet1')
Save_excel_sheet = Save_excel.get_sheet(0)
# 获取行数列数
Query_excel_sheet_nrows=Query_excel_sheet.nrows
Query_excel_sheet_nclos=Query_excel_sheet.ncols
# 获取头行
first_rows=Query_excel_sheet.row_values(0)
#结果列表
result=[]
result.append(first_rows)
for Query in range(1,Query_excel_sheet_nrows):
    Query_row = Query_excel_sheet.row_values(Query)
    txt=Query_row[0]
    ans_string=size(txt)
    Query_row[2]=ans_string
    result.append(Query_row)

for i in range(len(result)):
    for j in range(len(result[i])):
        Save_excel_sheet.write(i,j,result[i][j])
Save_excel.save('结果.xls')
print("finish")


