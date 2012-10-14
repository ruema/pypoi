'''
Created on 16.08.2012

@author: user
'''
import struct
import logging
import poi.functions
import re
from poi.functions import FUNCTION_MAP
from poi.utils import StringUtils

def col_by_name(colname):
    """'A' -> 0, 'Z' -> 25, 'AA' -> 26, etc
    """
    col=0
    for ch in colname:
        col=col*26+ord(ch)&0x1f
    return col-1

_re_cell_ex = re.compile(r"(\$?)([A-I]?[A-Z])(\$?)(\d+)", re.IGNORECASE)
def pack_rowcol(cell):
    """Convert an Excel cell reference string in A1 notation
    to numeric row/col notation.

    Returns: row, col_pack
    """
    m = _re_cell_ex.match(cell)
    if not m:
        raise Exception("Ill-formed single_cell reference: %s" % cell)
    col_abs, col, row_abs, row = m.groups()
    return int(row)-1, col_by_name(col)+(not row_abs and 0x4000)+(not col_abs and 0x8000)

def str_ref(row,col):
    co=col&0x3FFF
    if co:
        cc=''
        while co>0:
            co,c = divmod(co,26)
            cc=chr(65+c)+cc
    else:
        cc='A'
    return ('$%s$%s','%s$%s','$%s%s','%s%s')[col>>14]%(cc,row+1)

def _getvalue(worksheet,obj):
    if isinstance(obj,Ptg):
        return obj.getvalue(worksheet)
    return obj

class Reference(object):
    def __init__(self, worksheet, obj):
        self.worksheet=worksheet
        self.obj=obj

def _getref(worksheet,obj):
    if isinstance(obj,RefPtgBase):
        return Reference(worksheet,obj)
    return _getvalue(worksheet,obj)

class Ptg(object):
    """
    Ptg represents a syntactic token in a formula.  'PTG' is an acronym for
    'Parse ThinG.  Originally, the name referred to the single
    byte identifier at the start of the token, but in POI, <tt>Ptg</tt> encapsulates
    the whole formula token (initial byte + value data).
    
    Ptgs are logically arranged in a tree representing the structure of the
    parsed formula.  However, in BIFF files <tt>Ptg</tt>s are written/read in
    <em>Reverse-Polish Notation</em> order. The RPN ordering also simplifies formula
    evaluation logic, so POI mostly accesses <tt>Ptg</tt>s in the same way.
    """

    CLASS_REF = 0x00;
    CLASS_VALUE = 0x20;
    CLASS_ARRAY = 0x40;
    
    ptgClass = CLASS_REF

    @classmethod
    def read(cls, data, ofs, ofs2):
        raise NotImplementedError

    def mkstr(self, stack):
        stack.append(str(self))

    def calc(self, worksheet, stack):
        stack.append(self)
        
    def getvalue(self, worksheet):
        return None

class UnknownPtg(Ptg):
    pass


class ExpPtg(Ptg):
    @classmethod
    def read(cls, data, ofs, ofs2):
        self=cls()
        self._row,self._col = struct.unpack_from('<HH',data,ofs+1)
        return self, 5, 0
    
    def __str__(self):
        return "Exp"+str_ref(self._row,self._col)

    def getdata(self):
        return struct.pack('<BHH',0x01,self._row,self._col),None

class TblPtg(Ptg):
    pass


class ValueOperatorPtg(Ptg):
    instance=None
    
    @classmethod
    def read(cls, data, ofs, ofs2):
        if cls.instance is None:
            cls.instance=cls()
        return cls.instance,1,0
    
    def __str__(self):
        return self.op

    def mkstr(self, stack):
        op1=stack[-2]
        op2=stack[-1]
        stack[-2:]=[op1+self.op+op2]

    def calc(self, worksheet, stack):
        op1=_getvalue(worksheet, stack[-2])
        op2=_getvalue(worksheet, stack[-1])
        stack[-2:]=[self._calc(op1,op2)]

    def getdata(self):
        return chr(self.sid),None
    
class AddPtg(ValueOperatorPtg):
    sid=0x03
    op='+'
    
    def _calc(self, op1,op2):
        return op1+op2

class SubtractPtg(ValueOperatorPtg):
    sid=0x04
    op='-'

    def _calc(self, op1,op2):
        return op1-op2

class MultiplyPtg(ValueOperatorPtg):
    sid=0x05
    op='*'

    def _calc(self, op1,op2):
        return op1*op2

class DividePtg(ValueOperatorPtg):
    sid=0x06
    op='/'

    def _calc(self, op1,op2):
        return 1.0*op1/op2

class PowerPtg(ValueOperatorPtg):
    sid=0x07
    op='^'

    def _calc(self, op1,op2):
        return op1**op2

class ConcatPtg(ValueOperatorPtg):
    sid=0x08
    op='&'

    def _calc(self, op1,op2):
        return unicode(op1)+unicode(op2)

class LessThanPtg(ValueOperatorPtg):
    sid=0x09
    op='<'

    def _calc(self, op1,op2):
        return op1<op2
    
class LessEqualPtg(ValueOperatorPtg):
    sid=0x0A
    op='<='

    def _calc(self, op1,op2):
        return op1<=op2

class EqualPtg(ValueOperatorPtg):
    sid=0x0B
    op='='

    def _calc(self, op1,op2):
        return op1==op2

class GreaterEqualPtg(ValueOperatorPtg):
    sid=0x0C
    op='>='

    def _calc(self, op1,op2):
        return op1>=op2

class GreaterThanPtg(ValueOperatorPtg):
    sid=0x0D
    op='>'

    def _calc(self, op1,op2):
        return op1>op2

class NotEqualPtg(ValueOperatorPtg):
    sid=0x0E
    op='<>'

    def _calc(self, op1,op2):
        return op1!=op2

class IntersectionPtg(Ptg):
    pass


class UnionPtg(Ptg):
    pass


class RangePtg(Ptg):
    pass


class UnaryPlusPtg(ValueOperatorPtg):
    sid=0x12
    op='+...'

    def mkstr(self, stack):
        stack[-1]='+%s'%stack[-1]

    def calc(self, worksheet, stack):
        op1=_getvalue(worksheet, stack[-1])
        stack[-1]=op1

class UnaryMinusPtg(ValueOperatorPtg):
    sid=0x13
    op='-...'

    def mkstr(self, stack):
        stack[-1]='-%s'%stack[-1]

    def calc(self, worksheet, stack):
        op1=_getvalue(worksheet, stack[-1])
        stack[-1]=-op1

class PercentPtg(ValueOperatorPtg):
    sid=0x14
    op='%'

    def mkstr(self, stack):
        stack[-1]='%s%%'%stack[-1]

    def calc(self, worksheet, stack):
        op1=_getvalue(worksheet, stack[-1])
        stack[-1]=op1*0.01

class ParenthesisPtg(ValueOperatorPtg):
    sid=0x15
    op='(...)'

    def mkstr(self, stack):
        stack[-1]='(%s)'%stack[-1]

    def calc(self, worksheet, stack):
        pass

class MissingArgPtg(Ptg):
    instance=None
    
    @classmethod
    def read(cls, data, ofs, ofs2):
        if cls.instance is None:
            cls.instance=cls()
        return cls.instance,1,0
    
    def __str__(self):
        return ''

    def mkstr(self, stack):
        stack.append('')

    def getvalue(self, worksheet):
        return None

    def getdata(self):
        return '\x16',None
    

class ScalarConstantPtg(Ptg):
    def __init__(self, value):
        self.value=value
        
    def __str__(self):
        return str(self.value)

    def getvalue(self, worksheet):
        return self.value

class StringPtg(ScalarConstantPtg):
    @classmethod
    def read(cls, data, ofs, ofs2):
        val, ln = StringUtils.readString8B(data,ofs+1)
        return cls(val), ln+1,0

    def getdata(self):
        return '\x17'+StringUtils.writeString8B(self.value),None

class AttrPtg(Ptg):
    # flags 'volatile' and 'space', can be combined.
    # OOO spec says other combinations are theoretically possible but not likely to occur.
    semiVolatile = 0x01
    optiIf       = 0x02
    optiChoose   = 0x04
    optiSkip     = 0x08
    optiSum      = 0x10
    baxcel       = 0x20 # 'assignment-style formula in a macro sheet'
    space        = 0x40

    #SUM = AttrPtg(0x0010, 0, null, -1);

    # 00H = Spaces before the next token (not allowed before tParen token)
    SPACE_BEFORE = 0x00
    # 01H = Carriage returns before the next token (not allowed before tParen token)
    CR_BEFORE = 0x01
    # 02H = Spaces before opening parenthesis (only allowed before tParen token)
    SPACE_BEFORE_OPEN_PAREN = 0x02
    # 03H = Carriage returns before opening parenthesis (only allowed before tParen token)
    CR_BEFORE_OPEN_PAREN = 0x03
    # 04H = Spaces before closing parenthesis (only allowed before tParen, tFunc, and tFuncVar tokens)
    SPACE_BEFORE_CLOSE_PAREN = 0x04
    # 05H = Carriage returns before closing parenthesis (only allowed before tParen, tFunc, and tFuncVar tokens)
    CR_BEFORE_CLOSE_PAREN = 0x05
    # 06H = Spaces following the equality sign (only in macro sheets)
    SPACE_AFTER_EQUALITY = 0x06

    _jumpTable = None
    _chooseFuncOffset = -1

    @classmethod    
    def read(cls, data, ofs, ofs2):
        self = cls()
        self._options, self._data = struct.unpack_from('<BH',data,ofs+1)
        if self._options & cls.optiChoose:
            self._jumpTable = struct.unpack_from('<%dH'%(self._data+1),data,ofs+4)
            self._chooseFuncOffset = self._jumpTable.pop()
            return self, self._data*2+6, 0
        else:
            return self,4,0

    def __str__(self):
        opt = self._options
        if opt & self.semiVolatile:
            return "ATTR(semiVolatile)"
        if opt & self.optiIf:
            return "IF"
        if opt & self.optiChoose:
            return "CHOOSE"
        if opt & self.optiSkip:
            return ""
        if opt & self.optiSum:
            return "SUM"
        if opt & self.baxcel:
            return "ATTR(baxcel)"
        if opt & self.space:
            return " "*(self._data>>8)
        return "UNKNOWN ATTRIBUTE";

    def mkstr(self, stack):
        opt = self._options
        result = str(self)
        if opt & self.space:
            stack[-1]=result+stack[-1]
        elif opt & self.optiIf:
            stack[-1]='%s(%s)'%(result,stack[-1])
        elif opt & self.optiSkip:
            stack[-1]='%s%s'%(result,stack[-1]) # goto isn't a real formula element should not show up
        else:
            stack[-1]='%s(%s)'%(result,stack[-1])


class ErrPtg(ScalarConstantPtg):
    @classmethod
    def read(cls, data, ofs, ofs2):
        return cls(ord(data[ofs+1])),2,0

    def getdata(self):
        return '\x1C'+ord(self.value),None

class BoolPtg(ScalarConstantPtg):
    @classmethod
    def read(cls, data, ofs, ofs2):
        return cls(ord(data[ofs+1])!=0),2,0

    def getdata(self):
        return '\x1D\x01' if self.value else '\x1D\0',None

class IntPtg(ScalarConstantPtg):
    @classmethod
    def read(cls, data, ofs, ofs2):
        return cls(struct.unpack_from('<H',data,ofs+1)[0]),3,0

    def getdata(self):
        return struct.pack('<BH',0x1E,self.value),None

class NumberPtg(ScalarConstantPtg):
    @classmethod
    def read(cls, data, ofs, ofs2):
        return cls(struct.unpack_from('<d',data,ofs+1)[0]),9,0

    def getdata(self):
        return struct.pack('<Bd',0x1F,self.value),None

class ArrayPtg(Ptg):
    pass


class FuncPtg(Ptg):
    def __init__(self, func_num):
        self.func_num=func_num
        
    @classmethod
    def read(cls, data, ofs, ofs2):
        return cls(*struct.unpack_from('<H',data,ofs+1)), 3, 0
    
    def __str__(self):
        func=poi.functions.FUNCTION_TABLE[self.func_num]
        return "FUNC%d"%self.func_num if func is None else func.__name__

    def mkstr(self, stack):
        func=poi.functions.FUNCTION_TABLE[self.func_num]
        if func.minParams:
            params=stack[-func.minParams:]
            stack[-func.minParams:]=[]
        else:
            params=()
        stack.append('%s(%s)'%(func.__name__,';'.join(params)))

    def calc(self, worksheet, stack):
        func=poi.functions.FUNCTION_TABLE[self.func_num]
        if func.minParams:
            params=stack[-func.minParams:]
            stack[-func.minParams:]=[]
        else:
            params=()
        params=[_getvalue(worksheet,p) if t=='V' else _getref(worksheet, p) for t,p in zip(func.paramClasses,params)]
        stack.append(func(*params))

    def getdata(self):
        return struct.pack('<BH',0x21+self.ptgClass,self.func_num),None

class FuncVarPtg(Ptg):
    def __init__(self, num_args, func_num):
        self.num_args=num_args
        self.func_num=func_num
        
    @classmethod
    def read(cls, data, ofs, ofs2):
        return cls(*struct.unpack_from('<BH',data,ofs+1)), 4, 0
    
    def __str__(self):
        func=poi.functions.FUNCTION_TABLE[self.func_num]
        return "%s:%d"%(self.func_num if func is None else func.__name__,self.num_args)

    def mkstr(self, stack):
        func=poi.functions.FUNCTION_TABLE[self.func_num]
        if self.num_args:
            params=stack[-self.num_args:]
            stack[-self.num_args:]=[]
        else:
            params=()
        stack.append('%s(%s)'%(func.__name__,';'.join(params)))

    def calc(self, worksheet, stack):
        func=poi.functions.FUNCTION_TABLE[self.func_num]
        if self.num_args:
            params=stack[-self.num_args:]
            stack[-self.num_args:]=[]
        else:
            params=()
        params=[_getvalue(worksheet,p) if t=='V' else _getref(worksheet, p) for t,p in zip(func.paramClasses+func.paramClasses[-1]*len(params),params)]
        stack.append(func(*params))

    def getdata(self):
        return struct.pack('<BBH',0x22+self.ptgClass,self.num_args,self.func_num),None


class NamePtg(Ptg):
    @classmethod
    def read(cls, data, ofs, ofs2):
        self=cls()
        self._name_index,self._zero = struct.unpack_from('<HH',data,ofs+1)
        return self, 5, 0
    
    def __str__(self):
        return "Name%d"%self._name_index


class RefPtgBase(Ptg):
    def __init__(self, row, col):
        self._row=row
        self._col=col
        
    @classmethod
    def read(cls, data, ofs, ofs2):
        return cls(*struct.unpack_from('<HH',data,ofs+1)), 5, 0
    
    def __str__(self):
        return str_ref(self._row,self._col)
     
class RefPtg(RefPtgBase):
    def getvalue(self, worksheet):
        cell=worksheet.find_cell(self._row,self._col&0x3fff)
        if cell: return cell.get_value(worksheet)

    def getdata(self):
        return struct.pack('<BHH',0x24+self.ptgClass,self._row,self._col),None


class MemAreaPtg(Ptg):
    pass


class AreaPtg(Ptg):
    @classmethod
    def read(cls, data, ofs, ofs2):
        self=cls()
        self._firstrow,self._lastrow,self._firstcol,self._lastcol = struct.unpack_from('<4H',data,ofs+1)
        return self, 9, 0
    
    def __str__(self):
        return str_ref(self._firstrow,self._firstcol)+':'+str_ref(self._lastrow,self._lastcol)


class MemErrPtg(Ptg):
    pass


class MemFuncPtg(Ptg):
    pass


class RefErrorPtg(Ptg):
    pass


class AreaErrPtg(Ptg):
    pass


class RefNPtg(Ptg):
    pass


class AreaNPtg(Ptg):
    pass


class NameXPtg(Ptg):
    @classmethod
    def read(cls, data, ofs, ofs2):
        self = cls()
        self._sheetRefIndex, self._nameNumber, self._reserved = struct.unpack_from('<3H',data,ofs+1)
        return self, 7, 0

    def __str__(self, *args, **kwargs):
        return "[%d]%d!%d"%(self._sheetRefIndex,self._nameNumber,self._reserved)

class Ref3DPtg(Ptg):
    """
    Title:        Reference 3D Ptg
    Description:  Defined a cell in extern sheet.
    REFERENCE:
    """
    @classmethod
    def read(cls, data, ofs, ofs2):
        self = cls()
        self._extern_sheet_idx, self._row,self._col = struct.unpack_from('<3H',data,ofs+1)
        return self, 7, 0
    
    def __str__(self):
        return '%d!%s'%(self._extern_sheet_idx, str_ref(self._row,self._col))
        


class Area3DPtg(Ptg):
    """
    Title:        Area 3D Ptg - 3D reference (Sheet + Area)
    Description:  Defined a area in Extern Sheet.
    REFERENCE:
    """

    @classmethod
    def read(cls, data, ofs, ofs2):
        self = cls()
        (self.extern_sheet_idx,
         self.first_row, self.last_row,
         self.first_column, self.last_column) = struct.unpack_from('<5H',data,ofs+1)
        return self, 11,0
    
    def __str__(self):
        return "%d!%s:%s"%(self.extern_sheet_idx,str_ref(self.first_row,self.first_column),str_ref(self.last_row, self.last_column)) 

class DeletedRef3DPtg(Ptg):
    pass


class DeletedArea3DPtg(Area3DPtg):
    """
    Title:        Deleted Area 3D Ptg - 3D referecnce (Sheet + Area)
    Description:  Defined a area in Extern Sheet.
    REFERENCE:
    """
    
    @classmethod
    def read(cls, data, ofs, ofs2):
        self = cls()
        (self.extern_sheet_idx,
         self.first_row, self.last_row,
         self.first_column, self.last_column) = struct.unpack_from('<5H',data,ofs+1)
        return self, 11,0

    
    
LEX_RE=r'''(\$?[A-Z]+\$?[0-9+]|\$?R\d+\$?C\d+|\d*(?:\d\.|\.?\d)\d*(?:E[-+]?\d+)?|"(?:[^"]|"")*"|\w[.\w]*|'(?:[^']|'')*'|<>|<=|>=|[-+*/=<>:;,()&%^!])'''
LEX_REGEXP=re.compile(LEX_RE,re.IGNORECASE+re.LOCALE)
def FormulaLexer(formula):
    for tk in LEX_REGEXP.split(formula):
        if tk and tk.strip():
            if tk[0] not in ''''"''':
                tk=tk.upper()
            yield tk
            
OP_PTGS={
    '+':AddPtg,
    '-':SubtractPtg,
    '*':MultiplyPtg,
    '/':DividePtg,
    '^':PowerPtg,
    '&':ConcatPtg,
    '<':LessThanPtg,
    '<=':LessEqualPtg,
    '=':EqualPtg,
    '>':GreaterThanPtg,
    '>=':GreaterEqualPtg,
    '<>':NotEqualPtg,
    'u+':UnaryPlusPtg,
    'u-':UnaryMinusPtg,
    '%':PercentPtg,
}
_RVAdeltaRef =  {"R": 0, "V": 0x20, "A": 0x40, "D": 0x20}

class Formula(object):
    @classmethod
    def parse(cls, formula):
        self=cls()
        self.__toks=FormulaLexer(formula)
        self.next()
        self.ptgs=self.expr('V')
        print self.ptgs
        del self.__toks
        return self
    
    def next(self):
        try:
            self.curtok=next(self.__toks)
        except StopIteration:
            self.curtok=None
    
    def expr(self, argtype):
        result=self.prec0_expr(argtype)
        while self.curtok in ('=','<','>','<>','<=','>='):
            op=self.curtok;self.next()
            result+=self.prec0_expr(argtype)+[OP_PTGS[op]()]
        return result
    
    def prec0_expr(self, argtype):
        result=self.prec1_expr(argtype)
        while self.curtok in ('&',):
            op=self.curtok;self.next()
            result+=self.prec1_expr(argtype)+[OP_PTGS[op]()]
        return result
    
    def prec1_expr(self, argtype):
        result=self.prec2_expr(argtype)
        while self.curtok in ('+','-',):
            op=self.curtok;self.next()
            result+=self.prec2_expr(argtype)+[OP_PTGS[op]()]
        return result
    
    def prec2_expr(self, argtype):
        result=self.prec3_expr(argtype)
        while self.curtok in ('*','/'):
            op=self.curtok;self.next()
            result+=self.prec3_expr(argtype)+[OP_PTGS[op]()]
        return result
    
    def prec3_expr(self, argtype):
        result=self.prec4_expr(argtype)
        while self.curtok in ('^',):
            op=self.curtok;self.next()
            result+=self.prec4_expr(argtype)+[OP_PTGS[op]()]
        return result

    def prec4_expr(self, argtype):
        if self.curtok in ('+','-'):
            unary=[OP_PTGS['u'+self.curtok]()]
            self.next()
        else:
            unary=None
        result=self.primary(argtype)
        if self.curtok=='%':
            result+=[OP_PTGS['%']()]
            self.next()
        if unary:
            result+=unary
        return result

    def primary(self, argtype):
        if self.curtok in ('TRUE','FALSE'):
            result = [BoolPtg(self.curtok=='TRUE')]
            self.next()
        elif self.curtok[0]=='"':
            result = [StringPtg(self.curtok[1:-1].replace('""','"'))]
            self.next()
        elif self.curtok[0] in '.0123456789':
            val=float(self.curtok)
            result = [IntPtg(int(val)) if int(val)==val and val<65536 else NumberPtg(val)]
            self.next()
        elif self.curtok=='(':
            self.next()
            result = self.expr(argtype)+[ParenthesisPtg()]
            if self.curtok!=')':
                raise AssertionError('")" expected')
            self.next()
        else:
            name = self.curtok
            self.next()
            if self.curtok=='(':
                func=FUNCTION_MAP[name]
                cnt=0
                args=[]
                self.next()
                while self.curtok!=')':
                    if self.curtok in (',;'):
                        args.append(MissingArgPtg())
                    else:
                        args.extend(self.expr(func.paramClasses[len(args) if len(args)<len(func.paramClasses) else -1]))
                    cnt+=1
                    if self.curtok==')':
                        break
                    if self.curtok not in ',;':
                        raise Exception('";" expected')
                    self.next()
                self.next()
                # FUNC_if
                # FUNC_choose
                # FUNC
                if cnt>func.maxParams or cnt<func.minParams:
                    raise AssertionError('Wrong number of parameters %d for function %s'%(cnt,func.__name__))
                if func.maxParams==func.minParams:
                    result=args+[FuncPtg(func.index)]
                else:
                    result=args+[FuncVarPtg(cnt,func.index)]
            elif self.curtok==':':
                self.next()
                ref2=self.curtok
                self.next()
                result=['%s:%s'%(name,ref2)]
            else:
                #REF2D or Name
                ref=RefPtg(*pack_rowcol(name))
                ref.ptgClass=_RVAdeltaRef[argtype]
                result=[ref]
            # TODO: sheet
        return result
    
    
    
    
     
    @classmethod
    def read(cls, data, ofs, formulaLen=None):
        self = cls()
        if formulaLen is None:
            formulaLen = struct.unpack_from('<H',data,ofs)[0]
            ofs+=2
        ofs2=ofs+formulaLen
        pos=0
        ptgs=[]
        while pos<formulaLen:
            ptg, ln, ln2 = cls.createPtg(data,ofs+pos,ofs2)
            pos += ln
            ofs2 += ln2
            ptgs.append(ptg)
        if pos != formulaLen:
            logging.warning("Ptg array size mismatch %d/%d"%(pos,formulaLen))
        self.ptgs=ptgs
        return self

    def getdata(self):
        result=[]
        add=[]
        for ptg in self.ptgs:
            data,more = ptg.getdata()
            result.append(data)
            if more: add.append(more)
        result=''.join(result)
        add=''.join(add)
        return struct.pack('<H',len(result))+result+add
            
        
        
    def __str__(self):
        stack=[]
        for ptg in self.ptgs:
            ptg.mkstr(stack)
        if len(stack)>1: print '!!!!',stack[1:]
        return stack[0]

    def calc(self,worksheet):
        stack=[]
        for ptg in self.ptgs:
            ptg.calc(worksheet,stack)
        if len(stack)>1: print '!!!!',stack[1:]
        return _getvalue(worksheet,stack[0])

    BASE_PTG= [
        UnknownPtg,
        ExpPtg,    # 0x01
        TblPtg,    # 0x02
        AddPtg,    # 0x03
        SubtractPtg,    # 0x04
        MultiplyPtg,    # 0x05
        DividePtg,    # 0x06
        PowerPtg,    # 0x07
        ConcatPtg,    # 0x08
        LessThanPtg,    # 0x09
        LessEqualPtg,    # 0x0a
        EqualPtg,    # 0x0b
        GreaterEqualPtg,    # 0x0c
        GreaterThanPtg,    # 0x0d
        NotEqualPtg,    # 0x0e
        IntersectionPtg,    # 0x0f
        UnionPtg,    # 0x10
        RangePtg,    # 0x11
        UnaryPlusPtg,    # 0x12
        UnaryMinusPtg,    # 0x13
        PercentPtg,    # 0x14
        ParenthesisPtg,    # 0x15
        MissingArgPtg,    # 0x16
        StringPtg,    # 0x17
        None,
        AttrPtg,    # 0x19
        None,
        None,
        ErrPtg,    # 0x1c
        BoolPtg,    # 0x1d
        IntPtg,    # 0x1e
        NumberPtg,    # 0x1f
        ]

    CLASSIFIED_PTG= [
        ArrayPtg,    #0x20, 0x40, 0x60
        FuncPtg,    # 0x21, 0x41, 0x61
        FuncVarPtg,    #0x22, 0x42, 0x62
        NamePtg,    # 0x23, 0x43, 0x63
        RefPtg,    # 0x24, 0x44, 0x64
        AreaPtg,    # 0x25, 0x45, 0x65
        MemAreaPtg,    # 0x26, 0x46, 0x66
        MemErrPtg,    # 0x27, 0x47, 0x67
        None,
        MemFuncPtg,    # 0x29, 0x49, 0x69
        RefErrorPtg,    # 0x2a, 0x4a, 0x6a
        AreaErrPtg,    # 0x2b, 0x4b, 0x6b
        RefNPtg,    # 0x2c, 0x4c, 0x6c
        AreaNPtg,    # 0x2d, 0x4d, 0x6d
        None, None,
        None,None,None,None,None,None,None,None,
        None,
        NameXPtg,    # 0x39, 0x49, 0x79
        Ref3DPtg,    # 0x3a, 0x5a, 0x7a
        Area3DPtg,    # 0x3b, 0x5b, 0x7b
        DeletedRef3DPtg,    # 0x3c, 0x5c, 0x7c
        DeletedArea3DPtg,    # 0x3d, 0x5d, 0x7d
        None,None,
    ]

    @classmethod
    def createPtg(cls, data, ofs, ofs2):
        pid=ord(data[ofs])
        print hex(pid)
        if pid<0x20:
            result, ln, ln2 =cls.BASE_PTG[pid].read(data, ofs, ofs2)
        else:
            result, ln, ln2 = cls.CLASSIFIED_PTG[pid&0x1f].read(data, ofs, ofs2);
            if pid >= 0x60:
                result.ptgClass = Ptg.CLASS_ARRAY
            elif pid >= 0x40:
                result.ptgClass = Ptg.CLASS_VALUE
            else:
                result.ptgClass = Ptg.CLASS_REF
        print result
        return result, ln, ln2
    

