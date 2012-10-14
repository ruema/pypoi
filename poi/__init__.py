import weakref
from poi.excel import HSSFWorkbook, HSSFWorksheet, boolean_property
from poi.utils import BoundSheet, ExtendedFormat, FontRecord
from poi.formula import Formula

class Borders(object):
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

    def __init__(self, _format=None):
        self._format=_format or Format()
    
    Left=property(lambda self:self._format._get_border()&0xf,lambda self,val: self._format._set_border((val&0xf), 0xf))
    Right=property(lambda self:self._format._get_border()>>4&0xf,lambda self,val: self._format._set_border((val&0xf)<<4, 0xf0))
    Top=property(lambda self:self._format._get_border()>>8&0xf,lambda self,val: self._format._set_border((val&0xf)<<8, 0xf00))
    Bottom=property(lambda self:self._format._get_border()>>12&0xf,lambda self,val: self._format._set_border((val&0xf)<<12, 0xf000))
    DiagDown=property(lambda self:self._format._get_border2()>>14&0x1,lambda self,val: self._format._set_border((val&0x1)<<14, 0x4000))
    DiagUp=property(lambda self:self._format._get_border2()>>15&0x1,lambda self,val: self._format._set_border2((val&0x1)<<15, 0x8000))
    DiagStyle=property(lambda self:self._format._get_border3()>>21&0xf,lambda self,val: self._format._set_border3((val&0xf)<<21, 0x1e00000))

    LeftColor=property(lambda self:self._format._get_border2()&0x7f,lambda self,val: self._format._set_border2((val&0x7f), 0x007f))
    RightColor=property(lambda self:self._format._get_border2()>>7&0x7f,lambda self,val: self._format._set_border2((val&0x7f)<<7, 0x3f80))
    TopColor=property(lambda self:self._format._get_border3()&0x7f,lambda self,val: self._format._set_border3((val&0x7f), 0x007f))
    BottomColor=property(lambda self:self._format._get_border3()>>7&0x7f,lambda self,val: self._format._set_border3((val&0x7f)<<7, 0x3f80))
    DiagColor=property(lambda self:self._format._get_border3()>>14&0x7f,lambda self,val: self._format._set_border3((val&0x7f)<<14, 0x1fc000))


class Font(object):
    NONE             = 0;
    SUPER            = 1;
    SUB              = 2;
    SINGLE            = 1;
    DOUBLE            = 2;
    SINGLE_ACCOUNTING = 0x21;
    DOUBLE_ACCOUNTING = 0x22;
    
    def __init__(self, fnt=None, _format=None):
        self._fnt=fnt or FontRecord(0xc8,0,0x1fff,0x190,0,0,0,0,'Arial')
        self._format=_format

    def __set(self, **kw):
        self._fnt=self._fnt._replace(**kw)
        if self._format:
            self._format._Format__set(font_index=self._fnt)
    
    def __setattr(self,bit,val):
        self.__set(attributes=(self._fnt.attributes&~bit)|(bit if val else 0))

    Height=property(lambda self:self._fnt.height*0.05,lambda self, val:self.__set(height=int(val*20)))
    Bold=property(lambda self:self._fnt.bold_weight>0x280, lambda self, val: self.__set(bold_weight=0x190 if not val else 0x2BC))
    Weight=property(lambda self:self._fnt.bold_weight, lambda self, val: self.__set(bold_weight=val))
    Italic=property(lambda self:self._fnt.attributes&0x2!=0, lambda self, val: self.__setattr(0x2,val))
    StruckOut=property(lambda self:self._fnt.attributes&0x8!=0, lambda self, val: self.__setattr(0x8,val))
    Outline=property(lambda self:self._fnt.attributes&0x10!=0, lambda self, val: self.__setattr(0x10,val))
    Shadow=property(lambda self:self._fnt.attributes&0x20!=0, lambda self, val: self.__setattr(0x20,val))
    SuperSubScript=property(lambda self:self._fnt.super_sub_script, lambda self, val: self.__set(super_sub_script=val))
    Underline=property(lambda self:self._fnt.underline, lambda self, val: self.__set(underline=val))
    Fontname=property(lambda self:self._fnt.font_name, lambda self, val: self.__set(font_name=val))
    ColorIndex=property(lambda self:self._fnt.color_palette_index, lambda self, val: self.__set(color_palette_index=val))
    
class Format(object):
    
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
    
    def __init__(self, exformat=None, cell=None):
        self._format=exformat or ExtendedFormat(0,0,0,0,0,0,0,0,0x20c0)
        self._cell=cell

    def __set(self, **kw):
        self._format=self._format._replace(**kw)
        if self._cell:
            self._cell.xf_index=self._format

    def _get_border(self):
        return self._format.border_options
    
    def _set_border(self, val, mask):
        self.__set(border_options=self._get_border()&~mask|val)

    def _get_border2(self):
        return self._format.palette_options
    
    def _set_border2(self, val, mask):
        self.__set(palette_options=self._get_palette()&~mask|val)

    def _get_border3(self):
        return self._format.adtl_palette_options
    
    def _set_border3(self, val, mask):
        self.__set(adtl_palette_options=self._get_border3()&~mask|val)

    def _get_rotation(self):
        rot = self._format.alignment_options>>8&255
        if rot==0xff: return 0xff
        if rot>90:
            return 90-rot
        return rot
    
    def _set_rotation(self, rot):
        if rot==0xff:
            pass
        elif rot<0 and rot>=-90:
            rot=90-rot
        elif rot>=0 and rot<=90:
            pass
        else:
            # this is not allowed ;-)
            pass
        self.__set(alignment_options=self._format.alignment_options&~0xFF00|((rot&255)<<8))        
        
    def _get_font(self):
        fnt=self._format.font_index
        if not isinstance(fnt,FontRecord):
            fnt=self._cell._worksheet.parent.fonts[fnt] if self._cell else None
        return Font(fnt,self)
        
        
    Borders=property(lambda self:Borders(self), lambda self,value: self.__set(
            border_options=value._format._format.border_options,
            palette_options=value._format._format.palette_options,
            adtl_palette_options=(value._format._format.adtl_palette_options&0x3ffffff)|(self._format.adtl_palette_options&0xfc000000) ))
    Alignment=property(lambda self:self._format.alignment_options&7, lambda self, value:self.__set(alignment_options=self._format.alignment_options&~7|(value&7)))
    WrapText=property(lambda self:self._format.alignment_options&8!=0, lambda self, value:self.__set(alignment_options=self._format.alignment_options&~8|(value and 8 or 0)))
    VertAlign=property(lambda self:self._format.alignment_options>>4&7, lambda self, value:self.__set(alignment_options=self._format.alignment_options&~0x70|((value&7)<<4)))
    JustifyLast=property(lambda self:self._format.alignment_options&0x80!=0, lambda self, value:self.__set(alignment_options=self._format.alignment_options&~0x80|(value and 0x80 or 0)))
    Rotation=property(_get_rotation,_set_rotation)

    FillPattern=property(lambda self:self._get_border3()>>26&0x3f,lambda self,val: self._set_border3((val&0x3f)<<26, 0xfc000000))
    ForegroundColor=property(lambda self:self._format.fill_palette_options&0x7f, lambda self, value:self.__set(fill_palette_options=self._format.fill_palette_options&~0x7f|((value&0x7f))))
    BackgroundColor=property(lambda self:self._format.fill_palette_options>>7&0x7f, lambda self, value:self.__set(fill_palette_options=self._format.fill_palette_options&~0x3f80|((value&0x7f)<<7)))
    
    Indent=property(lambda self:self._format.indentation_options&0xf, lambda self, value:self.__set(indentation_options=self._format.indentation_options&~0xf|((value&0xf))))
    ShrinkToFit=property(lambda self:self._format.indentation_options>>4&0x1!=0, lambda self, value:self.__set(indentation_options=self._format.indentation_options&~0x10|((value&0x1)<<4)))
    MergeCells=property(lambda self:self._format.indentation_options>>5&0x1!=0, lambda self, value:self.__set(indentation_options=self._format.indentation_options&~0x20|((value&0x1)<<5)))
    ReadingOrder=property(lambda self:self._format.indentation_options>>6&0x3, lambda self, value:self.__set(indentation_options=self._format.indentation_options&~0xC0|((value&0x3)<<6)))
    #UserStyleName???

    Font=property(_get_font,lambda self,val:self.__set(font_index=val._fnt))
    
def validate_worksheet_name(name):
    if not 0<len(name)<32:
        raise AssertionError('Length of name must be <=31')
    for ch in '/\\?*[]:':
        if name in ch:
            raise AssertionError('Following characters are not allowed: / \\ ? * [ ] :')
    if name[0]=="'" or name[-1]=="'":
        raise AssertionError("Sheet names must not begin or end with (').")

class Range(object):
    def __init__(self, worksheet, row, column, rows, cols):
        self._worksheet=worksheet
        self._row=row
        self._column=column
        self._rows=rows
        self._cols=cols
        
    def get_value(self):
        try:
            row=self._worksheet.rows[self._row]
            cell=row.cells[self._column]
            return cell.get_value(self._worksheet)
        except KeyError:
            return None
        
    def set_value(self,value):
        cell=self._worksheet.get_cell(self._row, self._column)
        cell.set_value(value)
        
    def get_formula(self):
        try:
            row=self._worksheet.rows[self._row]
            cell=row.cells[self._column]
            return cell.get_formula(self._worksheet)
        except KeyError:
            return None
        
    def set_formula(self,value):
        formula=Formula.parse(value)
        cell=self._worksheet.get_cell(self._row, self._column)
        cell.set_formula(self._worksheet, formula)

    def get_numberformat(self):
        try:
            row=self._worksheet.rows[self._row]
            cell=row.cells[self._column]
            extformat=self._worksheet.parent.extendedformats[cell.xf_index]
            return self._worksheet.parent.numberformats_map[extformat.format_index]
        except:
            return None

    def get_format(self):
        #try:
        #    row=self._worksheet.rows[self._row]
        #    cell=row.cells[self._column]
        #    xfi=cell.xf_index
        #except KeyError:
        #    xfi=0
        cell=self._worksheet.get_cell(self._row, self._column)
        xfi=cell.xf_index
        if not isinstance(xfi,ExtendedFormat):
            xfi=self._worksheet.parent.extendedformats[xfi]
        return Format(xfi,cell)
    
    def set_format(self, form): 
        cell=self._worksheet.get_cell(self._row, self._column)
        cell.xf_index=form._format

    Value=property(get_value,set_value)
    Formula=property(get_formula,set_formula)
    NumberFormat=property(get_numberformat)
    Format=property(get_format,set_format)

__WORKSHEETS__=weakref.WeakKeyDictionary()
class Worksheet(object):
    def __new__(cls, worksheet):
        wr_result = __WORKSHEETS__.get(worksheet)
        result=wr_result and wr_result()
        if result is None:
            result=object.__new__(cls)
            result.__init__(worksheet)
            __WORKSHEETS__[worksheet]=weakref.ref(result)
        return result
    
    def __init__(self, worksheet):
        self._worksheet=worksheet
        
    def cells(self, row, column):
        return Range(self._worksheet, row, column, 1, 1)
    
    DefaultRowHeight = property(lambda self:self._worksheet.get_short(0x225), lambda self,val:self._worksheet.set_short(0x225,val))
    DefaultColWidth = property(lambda self:self._worksheet.get_short(0x55), lambda self,val:self._worksheet.set_short(0x55,val))
        
class Worksheets(object):
    def __init__(self, parent):
        self.workbook=parent
        
    def __getitem__(self, key):
        if isinstance(key,(int,long)):
            return Worksheet(self.workbook._workbook.sheets[key])
        for sh in self.workbook._workbook.sheets:
            if sh.sheetname==key:
                if sh.sheet is None:
                    sh.sheet=HSSFWorksheet(self.workbook._workbook, sh.sheetdata,sh.position_of_BOF)
                return Worksheet(sh.sheet)
        raise KeyError('Worksheet %r not found'%key)
    
    def __iter__(self):
        for sh in self.workbook._workbook.sheets:
            yield sh.sheetname

    def add(self, name):
        try:
            self[name]
            raise AssertionError('Workbook with name %r already exists!'%name)
        except KeyError:
            pass
        validate_worksheet_name(name)
        bsh=BoundSheet()
        bsh.sheetname=name
        bsh.sheet = sheet = HSSFWorksheet(self.workbook._workbook)
        self.workbook._workbook.sheets.append(bsh)
        return Worksheet(sheet)
        

class Workbook(object):
    def __init__(self, filename=None):
        self._workbook = HSSFWorkbook(filename)
        self._sheets = Worksheets(self)
      
    def save(self, filename):
        self._workbook.write(filename)
          
    sheets=property(lambda self:self._sheets)
    colors=property(lambda self:self._colors)
    styles=property(lambda self:self._styles)
    tablestyles=property(lambda self:self._tablestyles)
    
    date1904=property(lambda self:self._workbook.get_boolean(0x0022), lambda self,value:self._workbook.set_boolean(0x0022,value))
