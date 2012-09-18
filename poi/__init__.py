import weakref
from poi.excel import HSSFWorkbook

def validate_worksheet_name(name):
    if not 0<len(name)<32:
        raise AssertionError('Length of name must be <=31')
    for ch in '/\\?*[]:':
        if name in ch:
            raise AssertionError('Following characters are not allowed: / \\ ? * [ ] :')
    if name[0]=="'" or name[-1]=="'":
        raise AssertionError("Sheet names must not begin or end with (').")

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

class Worksheets(object):
    def __init__(self, parent):
        self.workbook=parent
        
    def __getitem__(self, key):
        if isinstance(key,(int,long)):
            return Worksheet(self.workbook._workbook.sheets[key])
        for sh in self.workbook._workbook.sheets:
            if sh.name==key:
                return Worksheet(sh)
        raise KeyError('Worksheet %r not found'%key)
    
    def __iter__(self):
        for sh in self.workbook._workbook.sheets:
            yield sh.name

    def add(self, name):
        try:
            self[name]
            raise AssertionError('Workbook with name %r already exists!'%name)
        except KeyError:
            pass
        validate_worksheet_name(name)
        sheet = HSSFWorkSheet(self.workbook._workbook, name)
        self.workbook._workbook.sheets.append(sheet)
        return sheet
        

class Workbook(object):
    def __init__(self, filename=None):
        self._workbook = HSSFWorkbook(filename)
        self._sheets = Worksheets(self)
      
    def write(self, filename):
        self._workbook.write(filename)  
    sheets=property(lambda self:self._sheets)
