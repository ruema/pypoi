'''
Created on 14.08.2012

@author: user
'''
import sys
from poi import HSSFWorkbook, Workbook


def dump_sheet(ws):
    print '-----',ws.name.encode('latin1','replace'),'-----'
    for row_num in xrange(max(ws.rows)+1 if ws.rows else 0):
        if row_num in ws.rows:
            cells = ws.rows[row_num].cells
            for col_num in xrange(max(cells)+1 if cells else 0):
                if col_num in cells:
                    val = cells[col_num].value
                    if isinstance(val,basestring):
                        print val.encode('latin1','replace'),
                    else:
                        print val
                print ';',
        print


wb=Workbook('/home/user/workspace/Queck/mso/poi-3.8/test-data/spreadsheet/DateFormats.xls')
for ws in wb._workbook.sheets:
    dump_sheet(ws)
print list(wb.sheets)
wb.sheets.add('Test')
print "Sheets:"
wb._workbook.write('test.xls')
#exit(0)

del wb

wb=HSSFWorkbook('test.xls')
print "Sheets:"
for ws in wb.sheets:
    dump_sheet(ws)

exit(0)

if __name__=='__main__':
    if len(sys.argv)==2:
        wb=HSSFWorkbook(sys.argv[1])
    else:
        #wb=HSSFWorkbook('/home/user/workspace/Queck/mso/poi-3.8/test-data/spreadsheet/unicodeNameRecord.xls')
        #wb=HSSFWorkbook('/home/user/workspace/Queck/mso/poi-3.8/test-data/spreadsheet/StringContinueRecords.xls')
        wb=HSSFWorkbook()
    wb.write('test.xls')
    for k,v in wb.numberformat.iteritems():
        print k,v
    for xf in wb.exformat:
        print xf.format_index, wb.numberformat.get(xf.format_index)
    exit(0)    
    for ws in wb.sheets:
        dump_sheet(ws)
    exit(0)

