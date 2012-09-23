'''
Created on 16.09.2012

@author: user
'''
from poi.poifs import CFBReader, CFBWriter
from poi.utils import record_stream, Record, BoundSheet, RecordList,\
    FontRecord, NumberFormat, RowInfo, CellInfo, StaticStrings, MulCellInfo,\
    ExtendedFormat, ColumnInfo, pack_short, pack_record
import poi.utils
import struct
import logging

class RecordContainer(object):
    
    def set_boolean(self, sid, value):
        record=self.urecord.get(sid)
        if record is None:
            record=Record(sid, None)
            self.urecord[sid]=record
            self.last_record.next=record
            self.last_record=record
        record.data=struct.pack('<3H',sid,2,1 if value else 0)

    def get_boolean(self, sid):
        record=self.urecord.get(sid)
        return struct.unpack_from('<H',record.data,4)[0] if record else 0

    def set_short(self, sid, value):
        record=self.urecord.get(sid)
        if record is None:
            record=Record(sid, None)
            self.urecord[sid]=record
            self.last_record.next=record
            self.last_record=record
        record.data=struct.pack('<3H',sid,2,value)

    def get_short(self, sid):
        record=self.urecord.get(sid)
        return struct.unpack_from('<H',record.data,4)[0] if record else 0

def boolean_property(sid):
    return property(lambda self:self.get_boolean(sid), lambda self,value:self.set_boolean(sid,value))        

std_format_strings = {
    # "std" == "standard for US English locale"
    # #### TODO ... a lot of work to tailor these to the user's locale.
    # See e.g. gnumeric-1.x.y/src/formats.c
    0x00: "General",
    0x01: "0",
    0x02: "0.00",
    0x03: "#,##0",
    0x04: "#,##0.00",
    0x05: "$#,##0_);($#,##0)",
    0x06: "$#,##0_);[Red]($#,##0)",
    0x07: "$#,##0.00_);($#,##0.00)",
    0x08: "$#,##0.00_);[Red]($#,##0.00)",
    0x09: "0%",
    0x0a: "0.00%",
    0x0b: "0.00E+00",
    0x0c: "# ?/?",
    0x0d: "# ??/??",
    0x0e: "m/d/yy",
    0x0f: "d-mmm-yy",
    0x10: "d-mmm",
    0x11: "mmm-yy",
    0x12: "h:mm AM/PM",
    0x13: "h:mm:ss AM/PM",
    0x14: "h:mm",
    0x15: "h:mm:ss",
    0x16: "m/d/yy h:mm",
    0x25: "#,##0_);(#,##0)",
    0x26: "#,##0_);[Red](#,##0)",
    0x27: "#,##0.00_);(#,##0.00)",
    0x28: "#,##0.00_);[Red](#,##0.00)",
    0x29: "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)",
    0x2a: "_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)",
    0x2b: "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)",
    0x2c: "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)",
    0x2d: "mm:ss",
    0x2e: "[h]:mm:ss",
    0x2f: "mm:ss.0",
    0x30: "##0.0E+0",
    0x31: "@",
    }


class HSSFWorkbook(RecordContainer):
    MAX_ROW = 0xFFFF
    MAX_COLUMN = 0x00FF

    # The maximum number of cell styles in a .xls workbook.
    # The 'official' limit is 4,000, but POI allows a slightly larger number.
    # This extra delta takes into account built-in styles that are automatically
    # created for new workbooks
    #
    # See http://office.microsoft.com/en-us/excel-help/excel-specifications-and-limits-HP005199291.aspx
    MAX_STYLES = 4030
    
    def __init__(self, filename=None):
        self.sheets=RecordList(BoundSheet)
        self.fonts=RecordList(FontRecord)
        self.numberformats=RecordList(NumberFormat)
        self.extendedformats=RecordList(ExtendedFormat)
        self.staticstrings=StaticStrings()
        if filename:
            self.read(filename)
            
    def write(self, filename):
        cfb = CFBWriter()
        for name in sorted(self.streams.iterkeys()):
            cfb.put(name,self.streams[name])
        cfb.put(('Workbook',),self.getdata())
        cfb.write(filename)
        
    def getdata(self):
        for sheet in self.sheets:
            if sheet.sheet:
                sheet.sheetdata=sheet.sheet.getdata()
        result=[]
        first=self.records.next
        sheetpos=-1
        while first:
            if sheetpos==-1 and first is self.sheets:
                sheetpos=len(result)
            result.append(first.data)
            first=first.next
        reslen=sum(map(len,result))
        result.append(self.staticstrings.getdata(reslen))
        result.append(struct.pack('<HH',0x000A,0))  #EOF
        reslen+=len(result[-2])+4
        
        for sheet in self.sheets:
            sheet.position_of_BOF=reslen
            result.append(sheet.sheetdata)
            reslen+=len(sheet.sheetdata)
        result[sheetpos]=self.sheets.data
        return ''.join(result)
        

    def read(self, filename):
        filehandle=open(filename)
        cfb = CFBReader(filehandle)
        self.streams = cfb.dirtree
        del cfb

        # Normally, the Workbook will be in a POIFS Stream
        # called "Workbook". However, some weird XLS generators use "WORKBOOK"
        workbook=None
        for wb in ('Book','BOOK','WORKBOOK','Workbook'):
            if (wb,) in self.streams:
                workbook = self.streams.pop((wb,))
        if not workbook:
            raise IOError('The file does not contain a Workbook-entry')

        loaders={
            0x0031: self.fonts.add,
            0x0085: self.sheets.add,
            0x00e0: self.extendedformats.add,
            0x00fc: self.staticstrings.read,            
            0x041E: self.numberformats.add,
        }

        urecord={}
        self.records=Record(0,0)
        last_record=self.records
        ofs=0        
        for sid, data in record_stream(workbook.data):
            if sid==0x000A: #EOF
                break
            new_record=loaders.get(sid,Record)(sid,data)
            if new_record:
                last_record.next=new_record
                last_record=new_record
                if sid not in urecord:
                    urecord[sid]=new_record
            #if new_record.__class__==Record:
            #    print '%04x(%08x): %s'%(sid,ofs,poi.utils.DEBUG_RECORDS.get(sid))
            ofs+=len(data)
        self.urecord=urecord
        self.last_record=last_record
        pos=len(workbook.data)
        for sheet in sorted(self.sheets,key=lambda s:s.position_of_BOF,reverse=True):
            sheet.sheetdata=workbook.data[sheet.position_of_BOF:pos]
            pos=sheet.position_of_BOF
        self.numberformats_map=dict(std_format_strings)
        self.numberformats_map.update(dict([(nf.index,nf.format) for nf in self.numberformats]))
        

class HSSFWorksheet(RecordContainer):
    def __init__(self, parent, data, ofs):
        self.columninfo=RecordList(ColumnInfo)
        self.parent=parent
        loaders={
            0x0006: self.add_cell, # Formula
            0x007d: self.columninfo.add,
            0x00bd: self.add_mulcell, # MulRKRecord
            0x00be: self.add_mulcell, # MulBlankRecord
            0x00fd: self.add_cell, #LabelSSTRecord
            0x0200: self.add_dimensions,
            0x0201: self.add_cell, #BlankRecord
            0x0203: self.add_cell, #NumberRecord
            0x0207: self.read_string, # StringRecord
            0x0208: self.add_row,
            0x027e: self.add_cell, #RKRecord
            
            0x00d7: Record.ignore, # DBCell
        }

        urecord={}
        self.rows={}
        self.records=Record(0,0)
        last_record=self.records        
        for sid, data in record_stream(data):
            if sid==0x000A: #EOF
                break
            new_record=loaders.get(sid,Record)(sid,data)
            if new_record:
                last_record.next=new_record
                last_record=new_record
                if sid not in urecord:
                    urecord[sid]=new_record
            #if 1 or new_record.__class__==Record:
            #    print '%04x(%08x): %s'%(sid,ofs,poi.utils.DEBUG_RECORDS.get(sid))
            ofs+=len(data)
        self.urecord=urecord

    def build_row(self, row):
        cols = row.cells.keys()
        cols.sort()
        self.firstCol=cols[0] if cols else 0
        self.lastCol=cols[-1] if cols else 0
        result=[]
        last_col=last_sid=-1
        for col in cols:
            cell=row.cells[col]
            cdata=cell.getdata(self,row.row_number,col)
            if last_col+1==col and last_sid==cell.sid and last_sid in (0x0201,0x27e):
                data=result[-1]
                if data[:2]=='\x01\x02':
                    data='\xbe\0'+data[2:]+'  '
                elif data[:2]=='\x7e\x02':
                    data='\xbd\0'+data[2:]+'  '
                cdata=data[4:-2]+cdata[8:]+struct.pack('<H',col)
                result[-1]=data[:2]+struct.pack('<H',len(cdata))+cdata
            else:
                result.append(cdata)
            last_col=col
            last_sid=cell.sid
        return ''.join(result)

    def build_rows(self):
        s_rows=[]
        s_cells=[]
        o_cells=[]
        rows=self.rows.keys()
        rows.sort()
        first_col=1000000
        last_col=0
        result=[]
        for crow in rows:
            row=self.rows[crow]
            row.row_number=crow
            s_cells.append(self.build_row(row))
            o_cells.append(len(s_cells[-1]))
            s_rows.append(row.data)
            first_col=min(first_col,row.firstCol)
            last_col=max(last_col,row.lastCol)
            if len(s_rows)==32 or crow==rows[-1]:
                result.extend(s_rows)
                result.extend(s_cells)
                ofs=len(s_rows)*20+sum(o_cells)
                result.append(pack_record(0x00d7,struct.pack('<i%dH'%len(s_rows),ofs,(len(s_rows)-1)*20,*o_cells[:-1])))
                s_rows=[]
                s_cells=[]
                o_cells=[]
                
        print rows[0],rows[-1],first_col,last_col
        return pack_short(0x0200,rows[0],rows[-1],first_col,last_col)+''.join(result)
        
    def getdata(self):
        result=[]
        first=self.records.next
        while first:
            if first is self.dimensions:
                result.append(self.build_rows())
            else:
                result.append(first.data)
            first=first.next
        result.append(struct.pack('<HH',0x000A,0))  #EOF
        return ''.join(result)

    def add_dimensions(self, sid, data):
        self.dimensions=Record(sid,data)
        return self.dimensions

    def add_row(self, sid, data):
        row=RowInfo(data)
        if row.row_number in self.rows:
            row.cells=self.rows[row].cells
        self.rows[row.row_number]=row
        return None
    
    def get_cell(self, row, col):
        if row not in self.rows:
            self.rows[row]=RowInfo()
        if col not in self.rows[row].cells:
            self.rows[row].cells[col]=CellInfo()
        return self.rows[row].cells[col]
        
    
    def add_cell(self, sid, data):
        cell, row, col=CellInfo.read(sid,data)
        if row not in self.rows:
            self.rows[row]=RowInfo()
        self.rows[row].cells[col]=cell
        self.lastcell=cell
        return None
    
    def add_mulcell(self, sid, data):
        cells, row, col=MulCellInfo.read(sid,data)
        if row not in self.rows:
            self.rows[row]=RowInfo()
        cls=self.rows[row].cells
        for c in cells:
            cls[col]=c
            col+=1
        return None
    
    def read_string(self, sid, data):
        # Stores the cached result of a text formula
        boundary, length=struct.unpack_from('<HH',data,2)
        boundary+=4
        ofs=6
        result=''
        while 1:
            result+=data[ofs+1:boundary].decode('utf-16le' if ord(data[ofs]) else 'latin1')
            if len(result)>=length:
                break
            ofs=boundary+4
            boundary=ofs+struct.unpack_from('<H',data,ofs-2)[0]
        if length!=len(result):
            logging.warning('String-Record-Length mismatch %d/%d'%(length,len(result)))
            result=result[:length]
        if self.lastcell.sid==0x0006:
            self.lastcell._value = result

    def add_string(self, value):
        return self.parent.staticstrings.string_map[value]