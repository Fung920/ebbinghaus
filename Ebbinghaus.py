#!/usr/bin/env python
#-*- encoding: utf-8 -*-
"""
# Author:        Fung Kong
# Email:         kyun.kong@gmail.com
# Created Time:  2017-07-27 23:52:01
# File Name:     Ebbinghaus.py
# Description:
    Generate the ebbinghaus memory curves for reciting English vocabs
    It will automatically output to excel if the xlwt package is installed,
        otherwise it will print the output to the terminal
#
"""
import datetime
# define the step of ebbinghaus
forgettingCurves=[0,1,2,4,7,15,30]
# define the number of the word list
wordList=26
# how much time will this cycle take
totalDay=wordList + forgettingCurves[6]
# List name
listName=' 'u'\u25a2'"List"
# beginning date
today=datetime.date.today()


# for exporting to Excel
import imp
try:
    imp.find_module('xlwt')
    import xlwt
    data=xlwt.Workbook()
    table=data.add_sheet('sheet1')
    style = xlwt.XFStyle()
    font=xlwt.Font()
    font.bold=True
    style.font=font

#write the Header
    table.write(0,0, 'Day', style)
    table.write(0,1, 'Date', style)
    table.write(0,2, 'FirstLearn', style)
    table.write(0,3, 'Review', style)

    for s in range(0, totalDay):
        tmp = ''
        for r in range(0, len(forgettingCurves))[::-1]:
            if s - forgettingCurves[r] >= 0 and s - forgettingCurves[r] < wordList \
                    and s >= forgettingCurves[r]:
                tmp = tmp + listName + str(s + 1 - forgettingCurves[r]).zfill(2) + ','
            if s < wordList:
                l = listName + str (s + 1).zfill(2)
            else:
                l = ''
        table.write(s+1, 0, str(s).zfill(2))
        table.write(s+1, 1, str(datetime.timedelta(days=s+1)+today))
        table.write(s+1, 2, l)
        table.write(s+1, 3, tmp[:-1])
    data.save('Ebbinghaus.xls')
except ImportError:
    #print the title
    print('%s%10s%20s%12s' % \
            ("Day", "Date", "FirstLearn", "Review"))

# print the output to the terminal if no xlwt package installed
    for s in range(0, totalDay):
        tmp = ''
        for r in range(0, len(forgettingCurves))[::-1]:
            if s - forgettingCurves[r] >= 0 and s - forgettingCurves[r] < wordList\
                    and s >=forgettingCurves[r]:
                        tmp = tmp + listName + str(s + 1 - \
                                forgettingCurves[r]).zfill(2) +','
        if s < wordList:
            l = listName + str(s+1).zfill(2)
        else:
            l = ''
        print('%s %15s %12s     %s' % (str(s).zfill(2), \
                datetime.timedelta(days=s+1)+today, l, tmp[:-1]))

# for reading the contents
#  wb = xlrd.open_workbook('Ebbinghaus.xls')
#  sh = wb.sheet_by_name(u'sheet1')
#  for rownum in range(sh.nrows):
    #  print (sh.row_values(rownum))

