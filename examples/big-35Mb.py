#!/usr/bin/env python
# tries stress SST, SAT and MSAT

from time import *
from poi import Workbook

#style = XFStyle()

wb = Workbook()
ws0 = wb.sheets.add('0')

colcount = 200 + 1
rowcount = 6000 + 1

t0 = time()

for col in xrange(colcount):
    for row in xrange(rowcount):
        ws0.cells(row, col).Value="BIG(%d, %d)" % (row, col)

t1 = time() - t0
print "\nsince starting elapsed %.2f s" % (t1)

print "Storing..."
wb.save('big-35Mb.xls')

t2 = time() - t0
print "since starting elapsed %.2f s" % (t2)


