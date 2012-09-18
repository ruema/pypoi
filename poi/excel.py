'''
Created on 16.09.2012

@author: user
'''
from poi.poifs import CFBReader, CFBWriter
from poi.utils import record_stream, Record, FontRecords
class HSSFWorkbook(object):
    MAX_ROW = 0xFFFF
    MAX_COLUMN = 0x00FF

    # The maximum number of cell styles in a .xls workbook.
    # The 'official' limit is 4,000, but POI allows a slightly larger number.
    # This extra delta takes into account built-in styles that are automatically
    # created for new workbooks
    #
    # See http://office.microsoft.com/en-us/excel-help/excel-specifications-and-limits-HP005199291.aspx
    MAX_STYLES = 4030
    
    fonts=None
    def __init__(self, filename=None):
        if filename:
            self.read(filename)
            
    def write(self, filename):
        cfb = CFBWriter()
        for name,entry in self.streams.iteritems():
            cfb.put(name,entry)
        cfb.put(('Workbook',),self.getdata())
        cfb.write(filename)
        
    def getdata(self):
        result=[]
        first=self.records.next
        while first:
            result.append(first.data)
            first=first.next
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
            0x0031: self.read_font,
        }

        urecord={}
        self.records=Record(0,0)
        last_record=self.records        
        for sid, data in record_stream(workbook.data):
            new_record=loaders.get(sid,Record)(sid,data)
            if new_record:
                last_record.next=new_record
                last_record=new_record
                if sid not in urecord:
                    urecord[sid]=new_record
            if sid==0x000A: #EOF
                break
        self.urecord=urecord
        
    def read_font(self, sid, data):
        fonts=self.fonts
        if fonts is None:
            self.fonts = fonts=FontRecords()
            fonts.add(data)
            return fonts
        fonts.add(data)
        return None