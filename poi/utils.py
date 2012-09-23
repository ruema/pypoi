'''
Created on 16.09.2012

@author: user
'''
import struct
import logging

class StringUtils(object):
    @staticmethod
    def readUnicodeLEString(data,ofs,chars):
        return data[ofs:ofs+2*chars].decode('utf-16le')
    
    @staticmethod
    def readCompressedUnicode(data, ofs, chars):
        return data[ofs:ofs+chars].decode('latin1')

    @staticmethod
    def readString(data, ofs, length=None):
        if length is None:
            length, is16bit = struct.unpack_from('<HB',data,ofs)
            add=3
        elif length is False:
            length, is16bit = struct.unpack_from('<BB',data,ofs)
            add=2
        else:
            is16bit = ord(data[ofs])
            add=1
        if is16bit:
            return data[ofs+add:ofs+add+length*2].decode('utf-16le'),add+length*2
        else:
            return data[ofs+add:ofs+add+length].decode('latin1'),add+length

    @staticmethod
    def readString8B(data, ofs):
        return StringUtils.readString(data, ofs, False)


    @staticmethod
    def writeString0(string):
        try:
            return 0, string.encode('latin-1')
        except:
            return 1, string.encode('utf-16le')

    @staticmethod
    def writeString(string):
        wide, data = StringUtils.writeString0(string)
        return struct.pack('<HB',len(string),wide)+data

    @staticmethod
    def writeString8B(string):
        wide, data = StringUtils.writeString0(string)
        return struct.pack('<BB',len(string),wide)+data


def record_stream(stream, ofs=0):
    sofs=ofs
    sid, length = struct.unpack_from('<HH',stream,ofs)
    while sid is not None:
        ofs+=length+4
        try:
            sid2, length = struct.unpack_from('<HH',stream,ofs)
        except: sid2=None
        if sid2 != 0x003C: # !Continue
            yield sid,stream[sofs:ofs]
            sofs=ofs
            sid=sid2

class Record(object):
    next=None
    def __init__(self, sid, data):
        self.sid=sid
        self.data=data
        
    @staticmethod
    def ignore(sid, data):
        return None

class RecordList(list, Record):
    def __init__(self, cls):
        self.cls=cls
        
    def add(self, sid, data):
        self.append(self.cls.read(data))
        return self if len(self)==1 else None

    @property
    def data(self):
        return ''.join(entry.getdata() for entry in self)

class FontRecord(object):
    SS_NONE             = 0;
    SS_SUPER            = 1;
    SS_SUB              = 2;
    U_NONE              = 0;
    U_SINGLE            = 1;
    U_DOUBLE            = 2;
    U_SINGLE_ACCOUNTING = 0x21;
    U_DOUBLE_ACCOUNTING = 0x22;
    # short height # in units of .05 of a point
    # short attributes
    # 0x01 - Reserved bit must be 0
    # 0x02 - is this font in italics
    # 0x04 - reserved bit must be 0
    # 0x08 - is this font has a line through the center
    # 0x10 - macoutline: some weird macintosh thing....but who understands those mac people anyhow
    # 0x20 - macshadow: some weird macintosh thing....but who understands those mac people anyhow
    # 7-6 - reserved bits must be 0
    # the rest is unused
    # short color_palette_index;
    # short bold_weight;
    # short super_sub_script;   // 00none/01super/02sub
    # byte  6_underline;          // 00none/01single/02double/21singleaccounting/22doubleaccounting
    # byte  family;             // ?? defined by windows api logfont structure?
    # byte  charset;            // ?? defined by windows api logfont structure?
    # byte  zero = 0;           // must be 0
    # possibly empty string never <code>null</code>
    # font_name;

    @classmethod
    def read(cls, data):
        self = cls()
        (self.height, self.attributes, self.color_palette_index, self.bold_weight, self.super_sub_script,
         self.underline, self.family, self.charset, _) = struct.unpack_from('<5H4B',data,4)
        self.font_name=StringUtils.readString8B(data, 18)[0]
        return self
    
    def getdata(self):
        fn = StringUtils.writeString8B(self.font_name)
        return struct.pack('<7H4B',0x0031,14+len(fn),self.height, self.attributes, 
            self.color_palette_index, self.bold_weight, self.super_sub_script,
            self.underline, self.family, self.charset, 0)+fn

class BoundSheet(object):
    sheet=None
    
    @classmethod
    def read(cls, data):
        """
        Title:        Bound Sheet Record (aka BundleSheet) (0x0085)
        Description:  Defines a sheet within a workbook.  Basically stores the sheet name
                      and tells where the Beginning of file record is within the HSSF
                      file.
        REFERENCE:  PG 291 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
        """
        self=cls()
        self.position_of_BOF, self.hidden = struct.unpack_from('<IH',data,4)
        self.sheetname = StringUtils.readString8B(data,10)[0]
        return self

    def getdata(self):
        name = StringUtils.writeString8B(self.sheetname)
        return struct.pack('<HHIH',0x0085,6+len(name),self.position_of_BOF,self.hidden)+name
            
class NumberFormat(object):
    @classmethod
    def read(cls, data):
        self=cls()
        self.index = struct.unpack_from('<H',data,4)[0]
        self.format=StringUtils.readString(data, 6)[0]
        return self

    def getdata(self):
        name = StringUtils.writeString(self.format)
        return struct.pack('<3H',0x041e,2+len(name),self.index)+name

class ExtendedFormat(object):
    NULL = 0xfff0 # null constant

    # xf type
    XF_STYLE            = 1;
    XF_CELL             = 0;

    # borders
    NONE                = 0x0;
    THIN                = 0x1;
    MEDIUM              = 0x2;
    DASHED              = 0x3;
    DOTTED              = 0x4;
    THICK               = 0x5;
    DOUBLE              = 0x6;
    HAIR                = 0x7;
    MEDIUM_DASHED       = 0x8;
    DASH_DOT            = 0x9;
    MEDIUM_DASH_DOT     = 0xA;
    DASH_DOT_DOT        = 0xB;
    MEDIUM_DASH_DOT_DOT = 0xC;
    SLANTED_DASH_DOT    = 0xD;
    
    # alignment
    GENERAL             = 0x0;
    LEFT                = 0x1;
    CENTER              = 0x2;
    RIGHT               = 0x3;
    FILL                = 0x4;
    JUSTIFY             = 0x5;
    CENTER_SELECTION    = 0x6;
    
    # vertical alignment
    VERTICAL_TOP        = 0x0;
    VERTICAL_CENTER     = 0x1;
    VERTICAL_BOTTOM     = 0x2;
    VERTICAL_JUSTIFY    = 0x3;
    
    # fill
    NO_FILL             = 0  ;
    SOLID_FILL          = 1  ;
    FINE_DOTS           = 2  ;
    ALT_BARS            = 3  ;
    SPARSE_DOTS         = 4  ;
    THICK_HORZ_BANDS    = 5  ;
    THICK_VERT_BANDS    = 6  ;
    THICK_BACKWARD_DIAG = 7  ;
    THICK_FORWARD_DIAG  = 8  ;
    BIG_SPOTS           = 9  ;
    BRICKS              = 10 ;
    THIN_HORZ_BANDS     = 11 ;
    THIN_VERT_BANDS     = 12 ;
    THIN_BACKWARD_DIAG  = 13 ;
    THIN_FORWARD_DIAG   = 14 ;
    SQUARES             = 15 ;
    DIAMONDS            = 16 ;
    
    @classmethod
    def read(cls, data):
        self=cls()
        (self.font_index, self.format_index, self.cell_options,
        self.alignment_options, self.indention_options,
        self.border_options, self.palette_options,
        self.adtl_palette_options, self.fill_palette_options) = struct.unpack_from('<7HIH',data,4)
        return self
    
    def getdata(self):
        return struct.pack('<9HIH',0x00E0,20,self.font_index, self.format_index, self.cell_options,
            self.alignment_options, self.indention_options,
            self.border_options, self.palette_options,
            self.adtl_palette_options, self.fill_palette_options)

def pack_record(sid, data):
    return struct.pack('<HH',sid,len(data))+data

def pack_short(sid, *value):
    return struct.pack('<HH%dH'%len(value),sid,len(value)*2,*value)

class ContinueWriter(object):
    RECORD_SIZE=8228
    CONTINUE = struct.pack('<H',0x003C)
    
    def __init__(self, sid):
        self.total_ofs = 4
        self.record_ofs = 4
        self.header_pos=1
        self.data = [struct.pack('<H',sid), '\0\0']
    
    def write_struct(self, format, *data): #@ReservedAssignment
        self.write(struct.pack(format,*data))
        
    def write(self, data):
        if self.record_ofs+len(data)>self.RECORD_SIZE:
            self.data[self.header_pos]=struct.pack('<H',self.record_ofs-4)
            self.data.append(self.CONTINUE)
            self.data.append('\0\0')
            self.total_ofs+=4
            self.record_ofs=4
            self.header_pos=len(self.data)-1
        self.data.append(data)
        self.record_ofs+=len(data)
        self.total_ofs+=len(data)
    
    @property    
    def available(self):
        return self.RECORD_SIZE-self.record_ofs
    
    def close(self):
        self.data[self.header_pos]=struct.pack('<H',self.record_ofs-4)
        return ''.join(self.data)


class StaticStrings(Record):
    def __init__(self):
        pass
    
    def read(self, sid, data):
        if sid=='0x00ff':
            return None
        
        next_border,self.num_strings,unique_strings = struct.unpack_from('<Hii',data,2)
        next_border+=4
        strings=[]
        pos=12
        for _i in xrange(unique_strings):
            assert pos+3<next_border
            nchars, options = struct.unpack_from('<HB', data,pos)
            pos += 3
            if options & 0x08: # richtext
                rtcount = struct.unpack_from('<H', data,pos)[0]
                pos += 2
            else: rtcount = 0
            if options & 0x04: # phonetic
                phosz = struct.unpack_from('<i', data,pos)[0]
                pos += 4
            else: phosz = 0
            accstrg = u''
            charsleft = nchars
            while 1:
                if options & 0x01:
                    # Uncompressed UTF-16
                    avail = min((next_border - pos) >> 1, charsleft)
                    accstrg += unicode(data[pos:pos+2*avail], "utf_16_le")
                    pos += 2*avail
                else:
                    avail = min(next_border - pos, charsleft)
                    accstrg += unicode(data[pos:pos+avail], 'latin_1')
                    pos += avail
                charsleft -= avail
                if charsleft == 0:
                    break
                next_border, options=struct.unpack_from('<HB',data,pos+2)
                next_border+=pos+4
                pos+=5
    
            if rtcount:
                runs = []
                rtcount*=2
                while 1:
                    avail = min((next_border - pos) >> 1, rtcount)
                    runs.extend(struct.unpack_from('<%dH'%avail,data,pos))
                    pos += 2*avail
                    rtcount-=avail
                    if rtcount==0:
                        break
                    pos+=4
                    next_border=pos+struct.unpack_from('<H',data,pos-2)[0]
            else:
                runs = ()
                
            if phosz:
                pho=''
                while 1:
                    avail = min(next_border - pos, phosz)
                    phosz+=data[pos:pos+avail]
                    pos += avail
                    phosz -= avail
                    if phosz==0:
                        break
                    pos+=4
                    next_border=pos+struct.unpack_from('<H',data,pos-2)[0]
            else:
                pho=None
            strings.append((accstrg,runs,pho))
        self.strings=strings
        self.string_map=dict([(s[0],i) for i,s in enumerate(self.strings)])
        return None
    
    def getdata(self,ofs):
        abs_rel_ofs=[]
        sst=ContinueWriter(0x00FC)
        sst.write(struct.pack('<ii',self.num_strings, len(self.strings)))
        for cnt, st in enumerate(self.strings):
            if cnt&7==0:
                abs_rel_ofs.append(ofs+sst.total_ofs)
                abs_rel_ofs.append(sst.record_ofs)

            string, runs, pho = st
            wide, cstr=StringUtils.writeString0(string)
            options=1 if wide else 0
            frm='<HB';xx=[]
            if runs: options|=4;frm+='H';xx.append(len(runs))
            if pho: options|=8;frm+='H';xx.append(len(pho))
            sst.write_struct(frm,len(string),options,*xx)
            ofs=0
            while ofs<len(cstr):
                if ofs>0:
                    sst.write('\x01' if wide else '\0')
                avail=sst.available
                if wide: avail&=~1
                sst.write(cstr[ofs:ofs+avail])
                ofs+=avail
            if runs:
                cstr=struct.pack('<%dH'%len(runs),*runs)
                ofs=0
                while ofs<len(cstr):
                    avail=sst.available&~3
                    sst.write(cstr[ofs:ofs+avail])
                    ofs+=avail
                sst.write(pho)
            
        return sst.close()+pack_record(0x00FF, struct.pack('<H%di'%len(abs_rel_ofs[:256]),8,*abs_rel_ofs[:256]))
    
class ColumnInfo(object):
    def __init__(self, *args):
        (self.firstCol, self.lastCol, self.colWidth,self.xfIndex,
         self.options, self.colInfo) = args

    @classmethod
    def read(cls, data):
        return cls(*struct.unpack_from('<6H',data,4))

    def getdata(self):
        return pack_short(0x007D,self.firstCol, self.lastCol, 
            self.colWidth,self.xfIndex, self.options, self.colInfo)

class RowInfo(Record):
    def __init__(self, data=None):
        (self.row_number, self.firstCol, self.lastCol, 
         self.height, self.optimize, self.reserved, self.options,
         self.xf_index) = struct.unpack_from('<8H',data or '\0'*20,4)
        self.cells={}
    
    @property
    def data(self):
        return pack_short(0x208,self.row_number, self.firstCol, self.lastCol, 
         self.height, self.optimize, self.reserved, self.options,
         self.xf_index)
        
def get_rkvalue(rk):
    if rk&2==2:
        val = rk>>2
    else:
        val = struct.unpack('<d',struct.pack('<ii',0,rk&~3))[0]
    if rk&1==1:
        val*=0.01
    return val

def set_rkvalue(rk):
    #return struct.pack('<d',rk)
    if int(rk)==rk and -0x20000000<=rk<=0x1fffffff:
        return struct.pack('<i',(int(rk)<<2)|2)
    d1=struct.pack('<d',rk)
    if d1[:4]=='\0\0\0\0' and ord(d1[4])&3==0:
        return d1[4:]
    rk*=100
    if int(rk)==rk and -0x20000000<=rk<=0x1fffffff:
        return struct.pack('<i',(int(rk)<<2)|3)
    d2=struct.pack('<d',rk)
    if d2[:4]=='\0\0\0\0' and ord(d2[4])&3==0:
        return chr(ord(d1[4])|1)+d1[5:]
    return d1

class ErrorCode(object):
    def __init__(self, errcode):
        self.errcode=errcode

class CellInfo(Record):
    _NOTSET=["NOT SET"]
    data=None
    formula=None
    def __init__(self, xf_index=0, value=None):
        self.xf_index=xf_index
        self._value=value
    
    @classmethod
    def read(self, sid, data):
        row,col,xf_index = struct.unpack_from('<3H',data,4)
        self=CellInfo(xf_index, self._NOTSET)
        self.sid=sid
        self.data=data[10:]
        if sid==0x0006:
            self.formula=self.data
        return self,row,col
        
    def get_value(self, worksheet):
        if self._value is self._NOTSET:
            if self.sid==0x0201:  # BlankRecord
                self._value=None
            elif self.sid==0x00FD: # LabelSSTRecord
                sst_index = struct.unpack_from('<i',self.data,0)[0]
                self._value=worksheet.parent.staticstrings.strings[sst_index][0]
            elif self.sid==0x027E: # RKRecord
                rk = struct.unpack_from('<i',self.data,0)[0]
                self._value=get_rkvalue(rk)
            elif self.sid==0x0006: # Formula
                value = struct.unpack_from('<HiH',self.data,0)
                if value[2]!=0xffff:
                    self._value = struct.unpack_from('<d',self.data,0)[0]
                elif value[0]==0: # STRING
                    self._value="<string %x>"%value[1]
                elif value[0]==1: # BOOLEAN
                    self._value=value[1]!=0
                elif value[0]==2: # ERROR_CODE
                    self._value=ErrorCode(value[1])
                elif value[0]==3: # EMPTY
                    self._value=None
                else:
                    logging.warning('Unknown SpecialCachedValueType %d'%value[0])
            elif self.sid==0x203:
                self._value=struct.unpack_from('<d',self.data,0)[0]
            else:
                raise AssertionError("Unknown cell type")
        return self._value
    
    def getdata(self, worksheet, row, col):
        more=''
        if self.data is not None and self.sid!=0x00fd:
            data=self.data
        elif self.formula:
            self.sid=0x0006
            if isinstance(self._value,(int,long,float)):
                data=struct.pack('<d',self._value)+self.formula[8:]
            elif isinstance(self._value,basestring):
                data=struct.pack('<HiH',0,0,0xffff)+self.formula[8:]
                sst=ContinueWriter(0x0207)
                wide, cstr=StringUtils.writeString0(self._value)
                sst.write_struct('<H',len(self._value))
                ofs=0
                while ofs<len(cstr):
                    sst.write('\x01' if wide else '\0')
                    avail=sst.available
                    if wide: avail&=~1
                    sst.write(cstr[ofs:ofs+avail])
                    ofs+=avail
                more=sst.close()
            elif self._value in (True, False):
                data=struct.pack('<HiH',1,int(self._value),0xffff)+self.formula[8:]
            elif isinstance(self._value,ErrorCode):
                data=struct.pack('<HiH',2,self._value.errcode,0xffff)+self.formula[8:]
            elif self._value is None:
                data=struct.pack('<HiH',3,0,0xffff)+self.formula[8:]
            else:
                raise AssertionError
        elif self._value is None:
            self.sid=0x0201
            data=''
        elif isinstance(self._value,(int,long,float)):
            data=set_rkvalue(self._value)
            self.sid=0x027E if len(data)==4 else 0x0203
        elif isinstance(self._value,basestring) or (self.data and self.sid==0x00fd):
            self.sid=0x00fd
            data=struct.pack('<i',worksheet.add_string(self.get_value(worksheet)))
        else:
            raise AssertionError('%04x:%r'%(self.sid,self._value))
        return struct.pack('<5H',self.sid,len(data)+6,row,col,self.xf_index)+data+more
        
    def set_value(self, value):
        self._value=value
        self.formula=None
        self.data=None

class MulCellInfo(Record):
    @classmethod
    def read(cls, sid, data):
        if sid==0x00be: # MulBlankRecord
            ccnt=(len(data)-6)/2
            idx = struct.unpack_from('<%dH'%ccnt,data,4)
            cells = [CellInfo(xf) for xf in idx[2:]]
        elif sid==0x00bd: # MulRKRecord
            idx = struct.unpack_from('<2H',data,4)
            xf_rk  = struct.unpack_from('<'+'Hi'*((len(data)-10)/6),data,8)
            cells=[CellInfo(xf_rk[i], get_rkvalue(xf_rk[i+1])) for i in xrange(0,len(xf_rk),2)]
        else:
            raise AssertionError('Unknown %04x'%sid)
        return cells, idx[0],idx[1] 
        
        
DEBUG_RECORDS={
0x0006:"FormulaRecord",
0x000A:"EOFRecord",
0x000C:"CalcCountRecord",
0x000D:"CalcModeRecord",
0x000E:"PrecisionRecord",
0x000f:"RefModeRecord",
0x0010:"DeltaRecord",
0x0011:"IterationRecord",
0x0012:"ProtectRecord",
0x0013:"PasswordRecord",
0x0014:"HeaderRecord",
0x0015:"FooterRecord",
0x0017:"ExternSheetRecord",
0x0018:"NameRecord",
0x0019:"WindowProtectRecord",
0x001A:"VerticalPageBreakRecord",
0x001B:"HorizontalPageBreakRecord",
0x001C:"NoteRecord",
0x001D:"SelectionRecord",
0x0022:"DateWindow1904Record",
0x0023:"ExternalNameRecord",
0x0026:"LeftMarginRecord",
0x0027:"RightMarginRecord",
0x0028:"TopMarginRecord",
0x0029:"BottomMarginRecord",
0x002a:"PrintHeadersRecord",
0x002b:"PrintGridlinesRecord",
0x002F:"FilePassRecord",
0x0031:"FontRecord",
0x003C:"ContinueRecord",
0x003d:"WindowOneRecord",
0x0040:"BackupRecord",
0x0041:"PaneRecord",
0x0042:"CodepageRecord",
0x0051:"DConRefRecord",
0x0055:"DefaultColWidthRecord",
0x0059:"CRNCountRecord",
0x005A:"CRNRecord",
0x005B:"FileSharingRecord",
0x005C:"WriteAccessRecord",
0x005D:"ObjRecord",
0x005E:"UncalcedRecord",
0x005f:"SaveRecalcRecord",
0x0063:"ObjectProtectRecord",
0x007D:"ColumnInfoRecord",
0x0080:"GutsRecord",
0x0081:"WSBoolRecord",
0x0082:"GridsetRecord",
0x0083:"HCenterRecord",
0x0084:"VCenterRecord",
0x0085:"BoundSheetRecord",
0x0086:"WriteProtectRecord",
0x008c:"CountryRecord",
0x008d:"HideObjRecord",
0x0092:"PaletteRecord",
0x009c:"FnGroupCountRecord",
0x009D:"AutoFilterInfoRecord",
0x00A0:"SCLRecord",
0x00A1:"PrintSetupRecord",
0x00BD:"MulRKRecord",
0x00BE:"MulBlankRecord",
0x00C1:"MMSRecord",
0x00D7:"DBCellRecord",
0x00DA:"BookBoolRecord",
0x00dd:"ScenarioProtectRecord",
0x00E0:"ExtendedFormatRecord",
0x00E1:"InterfaceHdrRecord",
0x00E2:"InterfaceEndRecord",
0x00E5:"MergeCellsRecord",
0x00EB:"DrawingGroupRecord",
0x00EC:"DrawingRecord",
0x00EC:"DrawingRecordForBiffViewer",
0x00ED:"DrawingSelectionRecord",
0x00FC:"SSTRecord",
0x00fd:"LabelSSTRecord",
0x00FF:"ExtSSTRecord",
0x013D:"TabIdRecord",
0x0160:"UseSelFSRecord",
0x0161:"DSFRecord",
0x01AA:"UserSViewBegin",
0x01AB:"UserSViewEnd",
0x01AE:"SupBookRecord",
0x01AF:"ProtectionRev4Record",
0x01B0:"CFHeaderRecord",
0x01B1:"CFRuleRecord",
0x01B2:"DVALRecord",
0x01B6:"TextObjectRecord",
0x01B7:"RefreshAllRecord",
0x01B8:"HyperlinkRecord",
0x01BC:"PasswordRev4Record",
0x01BE:"DVRecord",
0x01C1:"RecalcIdRecord",
0x0200:"DimensionsRecord",
0x0201:"BlankRecord",
0x0203:"NumberRecord",
0x0204:"LabelRecord",
0x0205:"BoolErrRecord",
0x0207:"StringRecord",
0x0208:"RowRecord",
0x020B:"IndexRecord",
0x0221:"ArrayRecord",
0x0225:"DefaultRowHeightRecord",
0x0236:"TableRecord",
0x023E:"WindowTwoRecord",
0x027E:"RKRecord",
0x0293:"StyleRecord",
0x041E:"FormatRecord",
0x04BC:"SharedFormulaRecord",
0x0809:"BOFRecord",
0x0867:"FeatHdrRecord",
0x0868:"FeatRecord",
0x088E:"TableStylesRecord",
0x0894:"NameCommentRecord",
0x089C:"HeaderFooterRecord",
}