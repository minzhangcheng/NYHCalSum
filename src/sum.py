#!/usr/bin/python

import sys
import xlrd
import xlwt

__detailFilename = '记账.xlsx'
__summaryFilename = '合计.xls'
__sheetName = {
    'detail': '明细',
    'summary': '合计',
    'type': '科目'
}
__symbol = {'支出': -1, '收入': 1}
__detailIndex = {
    1: 'year',
    4: 'type',
    6: 'value'
}
__detailIndex = {
    'year': 1,
    'type': 4,
    'value': 6
}
__translate = {
    'year': '年份',
    'type': '科目',
    'value': '金额',
    'summary': '合计',
    'begin': '期初数',
    'end': '年末结合数'
}

def readType(filename='', sheetname=''):
    if not filename:
        filename = __detailFilename
    if not sheetname:
        sheetname = __sheetName['type']
        
    rf = xlrd.open_workbook(filename)
    sh = rf.sheet_by_name(sheetname)
    type = dict()
    for i in range(0, sh.nrows):
        row = sh.row_values(i)
        t = row[0]
        s = row[1]        
        if __symbol[s] not in type:
            type.setdefault(__symbol[s], set())
        if t not in type[__symbol[s]]:
            type[__symbol[s]].add(t)
    for i in type:
        type[i] = list(type[i])
    return type

def readDetail(type=None, filename='', sheetname='', checkHeader=True, checkType=True):
    if not filename:
        filename = __detailFilename
    if not sheetname:
        sheetname = __sheetName['detail']
        
    rf = xlrd.open_workbook(filename)
    sh = rf.sheet_by_name(sheetname)
    if checkHeader:
        header = sh.row_values(0)
        for i in __detailIndex:
            if header[__detailIndex[i]] != __translate[i]:
                print('明细表格式不匹配，请确认明细表的格式如下：')
                cc = 'A'
                format = list()
                for i in __detailIndex:
                    format.append('第%s列为"%s"' 
                                  % (chr(ord(cc)+__detailIndex[i]), __translate[i]))
                print('\t%s' % '\n\t'.join(format))
                return False
    sum = dict()
    typeSet = set()
    if checkType:
        li = list()
        for i in type:
            li.extend(type[i])
        typeSet = set(li)
    for i in range(1, sh.nrows):
        row = sh.row_values(i)
        year = row[__detailIndex['year']]
        type = row[__detailIndex['type']]
        value = row[__detailIndex['value']]
        if year not in sum:
            sum.setdefault(year, dict())
        if checkType:
            if type not in typeSet:
                print('明细表第%d行中科目错误。' % i)
                return False
        if type not in sum[year]:
            sum[year].setdefault(type, 0)
        sum[year][type] += value
    return sum

def col(x):
    return chr(ord('A') + x)
    
def writeSummary(sum, type, filename='', sheetname=''):
    if not filename:
        filename = __summaryFilename
    if not sheetname:
        sheetname = __sheetName['summary']
    
    wf = xlwt.Workbook()
    sh = wf.add_sheet(__sheetName['summary'], cell_overwrite_ok=True)
    yearList = sorted(sum.keys())
    typeList = list()
    for i in type:
        typeList.extend(type[i])
    
    x = 0
    y = 0
    lastC = ''
    sh.write(y, x, __translate['year'])
    x += 1
    sh.write(y, x, __translate['begin'])
    x += 1
    for i in typeList:
        sh.write(y, x, i)
        x += 1
    sh.write(y, x, __translate['end'])
    lastC = col(x)
    
    for year in yearList:
        x = 0
        y += 1
        sh.write(y, x, year)
        x += 1
        if y > 1:
            formula = '%s%d' % (lastC, y)
            sh.write(y, x, xlwt.Formula(formula))
        x += 1
        formula = 'B%d' % (y + 1)
        for symbol in type:
            for t in type[symbol]:
                if t in sum[year]:
                    sh.write(y, x, sum[year][t])
                formula += '+(%d)*%s%d' % (symbol, col(x), y + 1)
                x += 1
        sh.write(y, x, xlwt.Formula(formula))
    
    x = 0
    y += 1
    sh.write(y, x, __translate['summary'])
    x += 2
    for i in typeList:
        formula = 'SUM(%s%d:%s%d)' % (col(x), 2, col(x), y)
        sh.write(y, x, xlwt.Formula(formula))
        x += 1
    wf.save(filename)
    
type = readType()
sum = readDetail(type)
if not sum:
    input('回车以退出程序。')
    sys.exit(2)
writeSummary(sum, type)
