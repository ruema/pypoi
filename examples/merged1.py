#!/usr/bin/env python
# -*- coding: windows-1251 -*-
# Copyright (C) 2005 Kiseliov Roman

from poi import Workbook, Borders, Format, Font

wb = Workbook()
ws0 = wb.sheets.add('sheet0')

fnt1 = Font()
fnt1.Fontname = 'Verdana'
fnt1.Bold = True
fnt1.Height = 18

brd1 = Borders()
brd1.Left = 0x06
brd1.Right = 0x06
brd1.Top = 0x06
brd1.Bottom = 0x06

fnt2 = Font()
fnt2.Fontname = 'Verdana'
fnt2.Bold = True
fnt2.Height = 14

brd2 = Borders()
brd2.Left = 0x01
brd2.Right = 0x01
brd2.Top = 0x01
brd2.Bottom = 0x01

fnt3 = Font()
fnt3.Fontname = 'Verdana'
fnt3.Bold = True
fnt3.Italic = True
fnt3.Height = 12

brd3 = Borders()
brd3.Left = 0x07
brd3.Right = 0x07
brd3.Top = 0x07
brd3.Bottom = 0x07

#fnt4 = Font()

style1 = Format()
style1.Font = fnt1
style1.Alignment = Format.CENTER
style1.VertAlign = Format.VERTICAL_BOTTOM
style1.FillPattern=Format.SOLID_FILL
style1.ForegroundColor=0x16
style1.Borders = brd1

style2 = Format()
style2.Font = fnt2
style2.Font.Italic=True
style2.Font.Underline=Font.DOUBLE_ACCOUNTING
style2.Font.Shadow=True
style2.Font.Weight=0x200
style2.Alignment = Format.RIGHT
style2.VertAlign = Format.VERTICAL_CENTER
style2.FillPattern=Format.SOLID_FILL
style2.ForegroundColor=0x1f
style2.Borders = brd2

style3 = Format()
style3.Font = fnt3
style3.Alignment = Format.LEFT
style3.VertAlign = Format.VERTICAL_TOP
style3.FillPattern=Format.SOLID_FILL
style3.ForegroundColor=0x24
style3.Borders = brd3
style3.Rotation=-90
style3.JustifyLast=True
style3.WrapText=True

#price_style = XFStyle()
#price_style.font = fnt4
#price_style.alignment = al2
#price_style.borders = brd3
#price_style.num_format_str = '_(#,##0.00_) "money"'
#
#ware_style = XFStyle()
#ware_style.font = fnt4
#ware_style.alignment = al3
#ware_style.borders = brd3

ws0.cells(3,3).Value="Hallo"
ws0.cells(3,3).Format=style1
ws0.cells(4,10).Value="Hallo"
ws0.cells(4,10).Format=style2
ws0.cells(14,16).Value="Hallo"
ws0.cells(14,16).Format=style3

#ws0.merge(3, 3, 1, 5, style1)
#ws0.merge(4, 10, 1, 6, style2)
#ws0.merge(14, 16, 1, 7, style3)
#ws0.col(1).width = 0x0d00


wb.save('merged1.xls')
