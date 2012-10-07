import weakref
from poi.excel import HSSFWorkbook, HSSFWorksheet, boolean_property
from poi.utils import BoundSheet
from poi.formula import Formula

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
        cell.set_formula(formula)
        cell.set_value(formula.calc(self._worksheet))

    def get_numberformat(self):
        try:
            row=self._worksheet.rows[self._row]
            cell=row.cells[self._column]
            extformat=self._worksheet.parent.extendedformats[cell.xf_index]
            return self._worksheet.parent.numberformats_map[extformat.format_index]
        except:
            return None

    Value=property(get_value,set_value)
    Formula=property(get_formula,set_formula)
    NumberFormat=property(get_numberformat)

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
