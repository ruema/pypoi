#!/usr/bin/env python
# -*- coding: windows-1251 -*-
# Copyright (C) 2005 Kiseliov Roman

from poi import Workbook, Borders, Font, Format

font0 = Font()
font0.Fontname = 'Times New Roman'
font0.StruckOut = True
font0.Bold = True

style0 = Format()
style0.Font = font0

wb = Workbook()
ws0 = wb.sheets.add('0')

ws0.cells(1, 1).Value='Test'
ws0.cells(1, 1).Format=style0

for i in range(0, 0x53):
    borders = Borders()
    borders.Left = i
    borders.Right = i
    borders.Top = i
    borders.Bottom = i

    ws0.cells(i,2).Format.Borders=borders
    ws0.cells(i,2).Format.Font=font0
    ws0.cells(i,3).Value=hex(i)
    ws0.cells(i,3).Format=ws0.cells(i,2).Format
    #ws0.write(i, 2, '', style)
    #ws0.write(i, 3, hex(i), style0)

#ws0.write_merge(5, 8, 6, 10, "")

wb.save('blanks.xls')
