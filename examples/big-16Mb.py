#!/usr/bin/env python
# tries stress SST, SAT and MSAT

import sys
from time import *
from poi import Workbook

#style = XFStyle()

wb = Workbook()
ws0 = wb.sheets.add('0')

colcount = 200 + 1
rowcount = 6000 + 1

t0 = time()
print "\nstart: %s" % ctime(t0)

print "Filling..."
for col in xrange(colcount):
    print "[%d]" % col,
    sys.stdout.flush() 
    for row in xrange(rowcount):
        #ws0.write(row, col, "BIG(%d, %d)" % (row, col))
        ws0.cells(row,col).Value = "BIG"

t1 = time() - t0
print "\nsince starting elapsed %.2f s" % (t1)

print "Storing..."
wb.save('big-16Mb.xls')

t2 = time() - t0
print "since starting elapsed %.2f s" % (t2)


