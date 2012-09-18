'''
Created on 16.09.2012

@author: user
'''
import struct

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
    while True:
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
    
    def write(self):
        fn = StringUtils.writeString8B(self.font_name)
        return struct.pack('<7H4B',0x0031,14+len(fn),self.height, self.attributes, 
            self.color_palette_index, self.bold_weight, self.super_sub_script,
            self.underline, self.family, self.charset, 0)+fn

class FontRecords(Record):
    def __init__(self):
        self.fonts=[]
    
    def add(self, data):
        self.fonts.append(FontRecord.read(data))
        
    @property
    def data(self):
        return ''.join(font.write() for font in self.fonts)
        