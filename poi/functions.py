import math
def sheetfunc(index,minParams,maxParams,returnClass,paramClasses, isVolatile=False, hasFootnote=False):
    def mk(func):
        func.index=index
        func.minParams=minParams
        func.maxParams=maxParams
        func.returnClass=returnClass
        func.paramClasses=paramClasses
        func.isVolatile=isVolatile
        func.hasFootnote=hasFootnote
        FUNCTION_TABLE[index]=func
        FUNCTION_MAP[func.__name__]=func
        return func
    return mk
FUNCTION_TABLE=[None]*400
FUNCTION_MAP={}

# Columns: (index, name, minParams, maxParams, returnClass, paramClasses, isVolatile, hasFootnote )

# Built-In Sheet Functions in BIFF2
@sheetfunc(0,0,30,'V','R')
def COUNT(*args):
    raise NotImplementedError

@sheetfunc(1,2,3,'V','VRR')
def IF(*args):
    raise NotImplementedError

@sheetfunc(2,1,1,'V','V')
def ISNA(val):
    raise NotImplementedError

@sheetfunc(3,1,1,'V','V')
def ISERROR(val):
    raise NotImplementedError

@sheetfunc(4,0,30,'V','R')
def SUM(*args):
    result = 0
    for v in args:
        if v is None:
            pass
        elif isinstance(v,basestring):
            pass
        elif hasattr(v,'worksheet'):
            result+=v.obj.getvalue(v.worksheet)
        else:
            result+=v
    return result

@sheetfunc(5,1,30,'V','R')
def AVERAGE(*args):
    raise NotImplementedError

@sheetfunc(6,1,30,'V','R')
def MIN(*args):
    raise NotImplementedError

@sheetfunc(7,1,30,'V','R')
def MAX(*args):
    raise NotImplementedError

@sheetfunc(8,0,1,'V','R')
def ROW(val):
    raise NotImplementedError

@sheetfunc(9,0,1,'V','R')
def COLUMN(*args):
    raise NotImplementedError

@sheetfunc(10,0,0,'V','')
def NA():
    raise NotImplementedError

@sheetfunc(11,2,30,'V','VR')
def NPV(*args):
    raise NotImplementedError

@sheetfunc(12,1,30,'V','R')
def STDEV(*args):
    raise NotImplementedError

@sheetfunc(13,1,2,'V','VV')
def DOLLAR(*args):
    raise NotImplementedError

@sheetfunc(14,2,2,'V','VV',hasFootnote=True)
def FIXED(*args):
    raise NotImplementedError

@sheetfunc(15,1,1,'V','V')
def SIN(val):
    return math.sin(val)

@sheetfunc(16,1,1,'V','V')
def COS(val):
    return math.cos(val)

@sheetfunc(17,1,1,'V','V')
def TAN(val):
    return math.tan(val)

@sheetfunc(18,1,1,'V','V')
def ATAN(val):
    return math.atan(val)

@sheetfunc(19,0,0,'V','')
def PI():
    return math.pi

@sheetfunc(20,1,1,'V','V')
def SQRT(val):
    return math.sqrt(val)

@sheetfunc(21,1,1,'V','V')
def EXP(val):
    return math.exp(val)

@sheetfunc(22,1,1,'V','V')
def LN(val):
    return math.log(val)

@sheetfunc(23,1,1,'V','V')
def LOG10(val):
    return math.log10(val)

@sheetfunc(24,1,1,'V','V')
def ABS(val):
    return abs(val)

@sheetfunc(25,1,1,'V','V')
def INT(val):
    return math.modf(val)[1]

@sheetfunc(26,1,1,'V','V')
def SIGN(val):
    return val.__cmp__(0)

@sheetfunc(27,2,2,'V','VV')
def ROUND(val,dgts=0):
    return round(val,dgts)

@sheetfunc(28,2,3,'V','VRR')
def LOOKUP(val):
    raise NotImplementedError


"""
29	INDEX	2	4	R	R V V V		
30	REPT	2	2	V	V V		
31	MID	3	3	V	V V V		
32	LEN	1	1	V	V		
33	VALUE	1	1	V	V		
34	TRUE	0	0	V	-		
35	FALSE	0	0	V	-		
36	AND	1	30	V	R		
37	OR	1	30	V	R		
38	NOT	1	1	V	V		
39	MOD	2	2	V	V V		
40	DCOUNT	3	3	V	R R R		
41	DSUM	3	3	V	R R R		
42	DAVERAGE	3	3	V	R R R		
43	DMIN	3	3	V	R R R		
44	DMAX	3	3	V	R R R		
45	DSTDEV	3	3	V	R R R		
46	VAR	1	30	V	R		
47	DVAR	3	3	V	R R R		
48	TEXT	2	2	V	V V		
49	LINEST	1	2	A	R R		x
50	TREND	1	3	A	R R R		x
51	LOGEST	1	2	A	R R		x
52	GROWTH	1	3	A	R R R		x
56	PV	3	5	V	V V V V V		
# Built-In Sheet Functions in BIFF2
57	FV	3	5	V	V V V V V		
58	NPER	3	5	V	V V V V V		
59	PMT	3	5	V	V V V V V		
60	RATE	3	6	V	V V V V V V		
61	MIRR	3	3	V	A V V		
62	IRR	1	2	V	A V		
63	RAND	0	0	V	-	x	
64	MATCH	2	3	V	V R R		
65	DATE	3	3	V	V V V		
66	TIME	3	3	V	V V V		
67	DAY	1	1	V	V		
68	MONTH	1	1	V	V		
69	YEAR	1	1	V	V		
70	WEEKDAY	1	1	V	V		x
71	HOUR	1	1	V	V		
72	MINUTE	1	1	V	V		
73	SECOND	1	1	V	V		
"""

@sheetfunc(74,0,0,'V','',isVolatile=True)
def NOW():
    raise NotImplementedError

"""
75	AREAS	1	1	V	R		
76	ROWS	1	1	V	A		
77	COLUMNS	1	1	V	A		
78	OFFSET	3	5	R	R V V V V	x	
82	SEARCH	2	3	V	V V V		
83	TRANSPOSE	1	1	A	A		
86	TYPE	1	1	V	V		
97	ATAN2	2	2	V	V V		
98	ASIN	1	1	V	V		
99	ACOS	1	1	V	V		
100	CHOOSE	2	30	R	V R		
101	HLOOKUP	3	3	V	V R R		x
102	VLOOKUP	3	3	V	V R R		x
105	ISREF	1	1	V	R		
109	LOG	1	2	V	V V		
111	CHAR	1	1	V	V		
112	LOWER	1	1	V	V		
113	UPPER	1	1	V	V		
114	PROPER	1	1	V	V		
115	LEFT	1	2	V	V V		
116	RIGHT	1	2	V	V V		
117	EXACT	2	2	V	V V		
118	TRIM	1	1	V	V		
119	REPLACE	4	4	V	V V V V		
120	SUBSTITUTE	3	4	V	V V V V		
121	CODE	1	1	V	V		
124	FIND	2	3	V	V V V		
125	CELL	1	2	V	V R	x	
126	ISERR	1	1	V	V		
127	ISTEXT	1	1	V	V		
128	ISNUMBER	1	1	V	V		
129	ISBLANK	1	1	V	V		
130	T	1	1	V	R		
131	N	1	1	V	R		
140	DATEVALUE	1	1	V	V		
141	TIMEVALUE	1	1	V	V		
142	SLN	3	3	V	V V V		
143	SYD	4	4	V	V V V V		
144	DDB	4	5	V	V V V V V		
148	INDIRECT	1	2	R	V V	x	
162	CLEAN	1	1	V	V		
163	MDETERM	1	1	V	A		
164	MINVERSE	1	1	A	A		
165	MMULT	2	2	A	A A		
167	IPMT	4	6	V	V V V V V V		
168	PPMT	4	6	V	V V V V V V		
169	COUNTA	0	30	V	R		
183	PRODUCT	0	30	V	R		
184	FACT	1	1	V	V		
189	DPRODUCT	3	3	V	R R R		
190	ISNONTEXT	1	1	V	V		
193	STDEVP	1	30	V	R		
194	VARP	1	30	V	R		
195	DSTDEVP	3	3	V	R R R		
196	DVARP	3	3	V	R R R		
197	TRUNC	1	1	V	V		x
198	ISLOGICAL	1	1	V	V		
199	DCOUNTA	3	3	V	R R R		
# New Built-In Sheet Functions in BIFF3
49	LINEST	1	4	A	R R V V		x
50	TREND	1	4	A	R R R V		x
51	LOGEST	1	4	A	R R V V		x
52	GROWTH	1	4	A	R R R V		x
197	TRUNC	1	2	V	V V		x
204	YEN	1	2	V	V V		x
205	FINDB	2	3	V	V V V		
206	SEARCHB	2	3	V	V V V		
207	REPLACEB	4	4	V	V V V V		
208	LEFTB	1	2	V	V V		
209	RIGHTB	1	2	V	V V		
210	MIDB	3	3	V	V V V		
211	LENB	1	1	V	V		
212	ROUNDUP	2	2	V	V V		
213	ROUNDDOWN	2	2	V	V V		
214	ASC	1	1	V	V		
215	JIS	1	1	V	V		x
219	ADDRESS	2	5	V	V V V V V		
220	DAYS360	2	2	V	V V		x
221	TODAY	0	0	V	-	x	
222	VDB	5	7	V	V V V V V V V		
227	MEDIAN	1	30	V	R ...		
228	SUMPRODUCT	1	30	V	A ...		
229	SINH	1	1	V	V		
230	COSH	1	1	V	V		
231	TANH	1	1	V	V		
232	ASINH	1	1	V	V		
233	ACOSH	1	1	V	V		
234	ATANH	1	1	V	V		
235	DGET	3	3	V	R R R		
244	INFO	1	1	V	V		
# New Built-In Sheet Functions in BIFF4
14	FIXED	2	3	V	V V V		x
204	USDOLLAR	1	2	V	V V		x
215	DBCS	1	1	V	V		x
216	RANK	2	3	V	V R V		
247	DB	4	5	V	V V V V V		
252	FREQUENCY	2	2	A	R R		
261	ERROR.TYPE	1	1	V	V		
269	AVEDEV	1	30	V	R ...		
270	BETADIST	3	5	V	V V V V V		
271	GAMMALN	1	1	V	V		
272	BETAINV	3	5	V	V V V V V		
273	BINOMDIST	4	4	V	V V V V		
274	CHIDIST	2	2	V	V V		
275	CHIINV	2	2	V	V V		
276	COMBIN	2	2	V	V V		
277	CONFIDENCE	3	3	V	V V V		
278	CRITBINOM	3	3	V	V V V		
279	EVEN	1	1	V	V		
280	EXPONDIST	3	3	V	V V V		
281	FDIST	3	3	V	V V V		
282	FINV	3	3	V	V V V		
283	FISHER	1	1	V	V		
284	FISHERINV	1	1	V	V		
285	FLOOR	2	2	V	V V		
286	GAMMADIST	4	4	V	V V V V		
287	GAMMAINV	3	3	V	V V V		
288	CEILING	2	2	V	V V		
289	HYPGEOMDIST	4	4	V	V V V V		
290	LOGNORMDIST	3	3	V	V V V		
291	LOGINV	3	3	V	V V V		
292	NEGBINOMDIST	3	3	V	V V V		
293	NORMDIST	4	4	V	V V V V		
294	NORMSDIST	1	1	V	V		
295	NORMINV	3	3	V	V V V		
296	NORMSINV	1	1	V	V		
297	STANDARDIZE	3	3	V	V V V		
298	ODD	1	1	V	V		
299	PERMUT	2	2	V	V V		
300	POISSON	3	3	V	V V V		
301	TDIST	3	3	V	V V V		
302	WEIBULL	4	4	V	V V V V		
303	SUMXMY2	2	2	V	A A		
304	SUMX2MY2	2	2	V	A A		
305	SUMX2PY2	2	2	V	A A		
306	CHITEST	2	2	V	A A		
307	CORREL	2	2	V	A A		
308	COVAR	2	2	V	A A		
309	FORECAST	3	3	V	V A A		
310	FTEST	2	2	V	A A		
311	INTERCEPT	2	2	V	A A		
312	PEARSON	2	2	V	A A		
313	RSQ	2	2	V	A A		
314	STEYX	2	2	V	A A		
315	SLOPE	2	2	V	A A		
316	TTEST	4	4	V	A A V V		
317	PROB	3	4	V	A A V V		
318	DEVSQ	1	30	V	R ...		
319	GEOMEAN	1	30	V	R ...		
320	HARMEAN	1	30	V	R ...		
321	SUMSQ	0	30	V	R ...		
322	KURT	1	30	V	R ...		
323	SKEW	1	30	V	R ...		
324	ZTEST	2	3	V	R V V		
325	LARGE	2	2	V	R V		
326	SMALL	2	2	V	R V		
327	QUARTILE	2	2	V	R V		
328	PERCENTILE	2	2	V	R V		
329	PERCENTRANK	2	3	V	R V V		
330	MODE	1	30	V	A		
331	TRIMMEAN	2	2	V	R V		
332	TINV	2	2	V	V V		
# New Built-In Sheet Functions in BIFF5
70	WEEKDAY	1	2	V	V V		x
101	HLOOKUP	3	4	V	V R R V		x
102	VLOOKUP	3	4	V	V R R V		x
220	DAYS360	2	3	V	V V V		x
336	CONCATENATE	0	30	V	V		
337	POWER	2	2	V	V V		
342	RADIANS	1	1	V	V		
343	DEGREES	1	1	V	V		
344	SUBTOTAL	2	30	V	V R		
345	SUMIF	2	3	V	R V R		
346	COUNTIF	2	2	V	R V		
347	COUNTBLANK	1	1	V	R		
350	ISPMT	4	4	V	V V V V		
351	DATEDIF	3	3	V	V V V		
352	DATESTRING	1	1	V	V		
353	NUMBERSTRING	2	2	V	V V		
354	ROMAN	1	2	V	V V		
# New Built-In Sheet Functions in BIFF8
358	GETPIVOTDATA	2	30				
359	HYPERLINK	1	2	V	V V		
360	PHONETIC	1	1	V	R		
361	AVERAGEA	1	30	V	R ...		
362	MAXA	1	30	V	R ...		
363	MINA	1	30	V	R ...		
364	STDEVPA	1	30	V	R ...		
365	VARPA	1	30	V	R ...		
366	STDEVA	1	30	V	R ...		
367	VARA	1	30	V	R ...		
"""