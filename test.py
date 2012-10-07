'''
Created on 14.08.2012

@author: user
'''
import sys
from poi import Workbook
import os

def dump_sheet(name,ws):
    print '-----',name.encode('latin1','replace'),'-----'
    for x in range(7):
        for y in range(3):
            print unicode(ws.cells(x,y).NumberFormat).encode('utf8'),unicode(ws.cells(x,y).Value).encode('utf8'),
            #ws.cells(x,y).Value="TT%i"%(x+10*y)
        print

def dump_supbooks(wb):
    for sb in wb._workbook.supbooks:
        if sb.tag:
            print sb.tag, sb.numsheets
        else:
            print sb.encoded_url
            for sw in sb.sheets:
                if isinstance(sw,basestring):
                    print sw
                else:
                    print sw.name
                    for r in sw.rows:
                        print r,sw.rows[r]



'''
dd=False
path='' if len(sys.argv)>1 else '/home/user/workspace/Queck/mso/poi-3.8/test-data/spreadsheet/'
for l in (sys.argv[1:] if len(sys.argv)>1 else os.listdir(path)):
    if l in ('51832.xls','50833.xls'):
        continue
    if l[-3:]=='xls':
        try:
            print l
            sys.stderr.write(l+'\n')
            wb=Workbook(path+l)
            for ws in wb.sheets:
                dump_sheet(ws,wb.sheets[ws])
            wb.save('test.xls')
            wb=Workbook('test.xls')
            for ws in wb.sheets:
                dump_sheet(ws,wb.sheets[ws])
        except int, e:
            print e

exit(0)
'''
                        
#wb=Workbook('/home/user/workspace/Queck/mso/poi-3.8/test-data/spreadsheet/StringContinueRecords.xls')
wb=Workbook('test.xls')
print wb.date1904
for ws in wb.sheets:
    dump_sheet(ws,wb.sheets[ws])
wb.save('test2.xls')
wb=Workbook('test2.xls')
dump_supbooks(wb)
#for ws in wb.sheets:
#    dump_sheet(ws,wb.sheets[ws])
exit(0)





    
    
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

