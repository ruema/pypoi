#!/usr/bin/env python
# -*- coding: windows-1251 -*-
# Copyright (C) 2005 Kiseliov Roman
from poi import Workbook

w = Workbook()
ws = w.sheets.add('F')

ws.cells(0, 0).Formula="-(1+1)"
ws.cells(1, 0).Formula="-(1+1)/(-2-2)"
ws.cells(2, 0).Formula="-(134.8780789+1)"
ws.cells(3, 0).Formula="-(134.8780789e-10+1)"
ws.cells(4, 0).Formula="-1/(1+1)+9344"

ws.cells(0, 1).Formula="-(1+1)"
ws.cells(1, 1).Formula="-(1+1)/(-2-2)"
ws.cells(2, 1).Formula="-(134.8780789+1)"
ws.cells(3, 1).Formula="-(134.8780789e-10+1)"
ws.cells(4, 1).Formula="-1/(1+1)+9344"

ws.cells(0, 2).Formula="A1*$B1"
ws.cells(1, 2).Formula="A2*B$2"
ws.cells(2, 2).Formula="A3*$B$3"
ws.cells(3, 2).Formula="A4*B4*sin(pi()/4)"
ws.cells(4, 2).Formula="A5%*B5*pi()/1000"

##############
## NOTE: parameters are separated by semicolon!!!
##############


ws.cells(5, 2).Formula="C1+C2+C3+C4+C5/(C1+C2+C3+C4/(C1+C2+C3+C4/(C1+C2+C3+C4)+C5)+C5)-20.3e-2"
ws.cells(5, 3).Formula="C1^2"
ws.cells(6, 2).Formula="SUM(C1;C2;;;;;C3;;;C4)"
#ws.cells(6, 3).Formula="SUM($A$1:$C$5)"

ws.cells(7, 0).Formula='"lkjljllkllkl"'
ws.cells(7, 1).Formula='"yuyiyiyiyi"'
ws.cells(7, 2).Formula='A8 & B8 & A8'
#ws.cells(8, 2).Formula='now()'

ws.cells(10, 2).Formula='TRUE'
ws.cells(11, 2).Formula='FALSE'
#ws.cells(12, 3).Formula='IF(A1>A2;3;"hkjhjkhk")'

w.save('formulas.xls')
