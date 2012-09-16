import struct
import os
import time
import logging

class PropertyEntry(object):
    def __init__(self, name, data=None, cid=''):
        self.name=name
        self.data=data
        self.size=len(data) if data is not None else 0 
        self.type=2 if data is not None else 1
        self.start=0
        self.left=-1
        self.right=-1
        self.child=-1
        self.cid=cid
        self.times=(0,0,0)
        
    @classmethod
    def read(cls, data, ofs):
        (name, namelen, type, _, left, right, child, cid, ct,mt,at,start,size)=struct.unpack_from('<64sHBBiii16siqqiq',data,ofs)
        name = name[:namelen-2].decode('utf-16le')
        result = cls(name, None, cid)
        result.left=left
        result.right=right
        result.child=child
        result.type=type
        result.start=start
        result.size=size
        result.times=(ct,mt,at)
        return result
        
    def getBytes(self):
        name=self.name[:31]
        return struct.pack('<64sHBBiii16siqqiq',name.encode('utf-16le'),len(name)*2+2,
            self.type,0,self.left,self.right,self.child,self.cid,self.times[0],self.times[1],self.times[2],self.start,
            self.size)

class CFBReader(object):
    _SIGNATURE = 0xE011CFD0, 0xE11AB1A1 # public static final long

    def __init__(self, filehandle):
        self.data=filehandle.read()
        if len(self.data)<512:
            raise IOError("Invalid file")

        (sig1,sig2,_cid,version,byteorder, self._blocksize, self._miniblocksize,
         _, propblocks, fatsectors, first_prop, _, self.minicutoff,
         first_minifat, num_minifat, first_difat, num_difat) = struct.unpack_from('<II16sIHHH6sIIIIIIIII',self.data,0)

        if (sig1,sig2)!=self._SIGNATURE:
            if sig1==0x04034b50:
                raise IOError("Invalid format: This is a Excel2007+ file")
            raise IOError("Invalid header signature; read %08x%08x, expected %08x%08x"%(sig1,sig2,self._SIGNATURE[0],self._SIGNATURE[1]))
        if version not in (0x0003003B, 0x0003003E, 0x0004003E):
            raise IOError("Invalid version %08x"%version)
        if byteorder!=0xFFFE:
            raise IOError("Invalid byteorder %04x"%byteorder)

        fatpersec=1<<(self._blocksize-2)

        #read DIFAT
        difat = list(struct.unpack_from('<109I',self.data,19*4))
        while first_difat<0xfffffffe and num_difat>0:
            difat+=list(struct.unpack_from('<%dI'%fatpersec,self.data,(first_difat+1)<<self._blocksize))
            num_difat-=1
            first_difat=difat.pop()
        if first_difat<0xfffffffe or num_difat!=0:
            raise IOError("corrupt difat-chain %04x/%d"%(first_difat,num_difat))
        
        #read fat
        self._fat=[]
        for sec in difat[:fatsectors]:
            self._fat+=list(struct.unpack_from('<%dI'%fatpersec,self.data,(sec+1)<<self._blocksize))
        for sec in difat[fatsectors:]:
            if sec!=0xffffffff:
                raise IOError("Used fatsector beyond limit")
        
        #read minifat
        self._minifat=[]
        chain = list(self._fat_chain_iterator(first_minifat))
        if len(chain)!=num_minifat:
            logging.warning("corrupt minifat-chain %d/%d"%(len(chain),num_minifat))
        for sec in chain:
            self._minifat+=list(struct.unpack_from('<%dI'%fatpersec,self.data,(sec+1)<<self._blocksize))

        # read properties
        proppersec=1<<(self._blocksize-7)
        self._properties=[]
        chain = list(self._fat_chain_iterator(first_prop))
        if propblocks and len(chain)*proppersec!=propblocks:
            print len(chain)*proppersec,propblocks
            raise IOError("corrupt property-chain")
        for sec in chain:
            idx=(sec+1)<<self._blocksize
            for _ in xrange(proppersec):
                self._properties.append(PropertyEntry.read(self.data,idx))
                idx+=128
        self._ministream=self._get_stream(self._properties[0].start)
        for prop in self._properties[1:]:
            if prop.type==0: continue
            if prop.type==2:
                if prop.size>=self.minicutoff:
                    prop.data=self._get_stream(prop.start)
                else:
                    prop.data=self._get_ministream(prop.start)
        self.dirtree={}
        self._walk_dirs(self._properties[0].child)
        del self.data # finished reading

    def _fat_chain_iterator(self, index):
        while index<0xFFFFFFFE:
            yield index
            index = self._fat[index]

    def _minifat_chain_iterator(self, index):
        while index<0xFFFFFFFE:
            yield index
            index = self._minifat[index]

    def _get_stream(self, index):
        result=[]
        for sec in self._fat_chain_iterator(index):
            idx=(sec+1)<<self._blocksize
            result.append(self.data[idx:idx+(1<<self._blocksize)])
        return ''.join(result)

    def _get_ministream(self, index):
        result=[]
        for sec in self._minifat_chain_iterator(index):
            idx=sec<<self._miniblocksize
            result.append(self._ministream[idx:idx+(1<<self._miniblocksize)])
        return ''.join(result)

    def _walk_dirs(self, index, parents=()):
        entry = self._properties[index]
        name = parents+(entry.name,)
        #if entry.type==2:
        self.dirtree[name]=entry
        if entry.left!=-1:
            self._walk_dirs(entry.left, parents)
        if entry.right!=-1:
            self._walk_dirs(entry.right, parents)
        if entry.child!=-1:
            self._walk_dirs(entry.child, name)

class CFBWriter(object):
    
    def __init__(self, blocksize=9, miniblocksize=6):
        self._blocksize = blocksize
        self._miniblocksize = miniblocksize
        self.minicutoff = 4096
        self.dirtree=PropertyEntry('Root Entry', None, '\x20\x08\x02\0\0\0\0\0\xc0\0\0\0\0\0\0F')
        self.dirtree.child={}
        
    def put(self, direntry, data, cid=''):
        new = None
        cur = self.dirtree
        for name in direntry:
            if name in cur.child:
                cur=cur.child[name]
            else:
                new=PropertyEntry(name)
                new.child={}
                cur.child[name]=new
                cur=new
        if new is None: # direntry already exists
            if cur.type==2 or cur.child:
                raise AssertionError("Stream already exists: /%s"%'/'.join(direntry))
        
        if isinstance(data, PropertyEntry):
            cur.data=data.data
            cur.cid=data.cid
            cur.size=data.size
            cur.type=data.type
            cur.times=data.times
        else:
            cur.data=data
            cur.size=len(data) if data else 0
            cur.type=2 if data is not None else 1
        if cur.type==2:
            cur.child=-1
        
    def write(self, filename):
        self.__makeproperties()
        self.__makesmallstream()
        propblocks = (len(self.properties)*128+(1<<self._blocksize)-1)>>self._blocksize
        self.totalblocks+= propblocks
        while True:
            # fat sectors
            fatsectors = (self.totalblocks*4+(1<<self._blocksize)-1)>>self._blocksize
            # difat sectors
            difatsectors = 0
            rest = fatsectors-109
            while rest>0:
                rest -= (1<<(self._blocksize-2))-1
                difatsectors += 1
            self.totalblocks+=fatsectors + difatsectors
            if fatsectors == (self.totalblocks*4+(1<<self._blocksize)-1)>>self._blocksize:
                break
        self.__makefat(difatsectors,fatsectors,self.sbatfatsize, propblocks, self.smallblocks)
        self.properties[0].start = difatsectors+fatsectors+self.sbatfatsize+propblocks

        # Write Header
        version=3 if self._blocksize==9 else 4
        SIGNATURE = 0xE011CFD0, 0xE11AB1A1 # public static final long
        header=struct.pack('<II16sIHHH6sIIIIIIIII',SIGNATURE[0],SIGNATURE[1],'\0',
            0x0003003E if version==3 else 0x0004003E,
            0xFFFE, self._blocksize,self._miniblocksize,'\0',
            0 if version==3 else propblocks, fatsectors,
            difatsectors+fatsectors+self.sbatfatsize,
            0,self.minicutoff,difatsectors+fatsectors if self.sbatfatsize else 0xfffffffe,
            self.sbatfatsize,0 if difatsectors else 0xfffffffe,difatsectors)
        
        intsperblock=1<<(self._blocksize-2)
        df=range(difatsectors,difatsectors+fatsectors)+[0xffffffff]*intsperblock
        header+=struct.pack('<109I',*df[:109])
        if self._blocksize!=9:
            header+='\0\0\0\0'*(intsperblock-128)
        filehandle=open(filename,'wb')
        filehandle.write(header)
        # Write DIFAT
        if difatsectors:
            ofs=109
            for cnt in xrange(difatsectors):
                filehandle.write(struct.pack('<%sI'%intsperblock,*(df[ofs:ofs+intsperblock-1]+[cnt+1])))
        # Write XBAT
        cnt = fatsectors * intsperblock
        if len(self.XBATdata)<cnt: self.XBATdata.extend([0xffffffff]*(cnt-len(self.XBATdata)))
        filehandle.write(struct.pack('<%sI'%cnt,*self.XBATdata))
        # Write SBAT
        cnt = self.sbatfatsize * intsperblock
        sBATblocks=len(self.SBATdata)
        missing = (cnt-sBATblocks)
        if sBATblocks<cnt: self.SBATdata.extend([0xffffffff]*missing)
        filehandle.write(struct.pack('<%sI'%cnt,*self.SBATdata))
        # Write Properties
        data = ''.join(map(lambda x:x.getBytes(),self.properties))
        filehandle.write(data+'\0'*(propblocks*intsperblock*4-len(data)))
        # Write SmallStream
        blocksize=1<<self._miniblocksize
        for data in self.SBATstreams:
            filehandle.write(data)
            if len(data) & (blocksize-1):
                filehandle.write('\0'*(blocksize-(len(data) & (blocksize-1))))
        used = (sBATblocks*blocksize)&((1<<self._blocksize)-1)
        if used:
            filehandle.write('\0'*((1<<self._blocksize)-used))
        # Write BigStream
        blocksize=1<<self._blocksize
        for data in self.XBATstreams:
            filehandle.write(data)
            if len(data) & (blocksize-1):
                filehandle.write('\0'*(blocksize-(len(data) & (blocksize-1))))
        filehandle.close()
        
    def __appendXBAT(self, offset, length):
        self.XBATdata.extend(xrange(offset+1,offset+length))
        self.XBATdata.append(0xFFFFFFFE)
        return offset + length
        
    def __makefat(self, difatsectors, fatsectors, sbatfatsize, propblocks, smallblocks):
        self.XBATdata=[0xFFFFFFFC]*difatsectors + [0xFFFFFFFD]*fatsectors
        offset = difatsectors + fatsectors
        if sbatfatsize>0: offset = self.__appendXBAT(offset, sbatfatsize)
        if propblocks>0: offset = self.__appendXBAT(offset, propblocks)
        if smallblocks>0: offset = self.__appendXBAT(offset, smallblocks)
            
        self.XBATstreams=[]
        for entry in self.properties[1:]:
            if entry.size<self.minicutoff:
                pass
            else:
                entry.start = offset
                self.XBATstreams.append(entry.data)
                blocks=(entry.size+(1<<self._blocksize)-1)>>self._blocksize
                offset = self.__appendXBAT(offset, blocks)
        
    @staticmethod
    def __propcompare(name1, name2):    
        result = len(name1) - len(name2) #  int
        if result == 0:
            if name1=="_VBA_PROJECT":
                result = 1
            elif name2=="_VBA_PROJECT":
                result = -1
                #_VBA_PROJECT, it seems, will always come last
            else:
                result = cmp(name1.startswith("__"),name2.startswith("__"))
                #Betweeen __SRP_0 and __SRP_1 just sort as normal
                if result == 0:
                    result = cmp(name1.lower(),name2.lower())
                    #The default case is to sort names ignoring case
        return result
    
    def __makeproperties(self):
        self.properties=[]
        entry = self.__makeproperty(self.dirtree)
        entry.type=5 # root storage type

    def __makeproperty(self, entry):
        self.properties.append(entry)
        if isinstance(entry.child,dict):
            data = entry.child
            children = sorted(data, cmp=self.__propcompare)
            midpoint = len(children) / 2
            name = children[midpoint]
            entry.child = len(self.properties)
            chentry = entry2 = self.__makeproperty(data[name])
            for j in xrange(midpoint-1,-1,-1):
                name = children[j]
                entry2.left = len(self.properties)
                entry2 = self.__makeproperty(data[name])
            entry2 = chentry
            for j in xrange(midpoint + 1,len(children)):
                name = children[j]
                entry2.right = len(self.properties)
                entry2 = self.__makeproperty(data[name])
        return entry
            
    def __makesmallstream(self):
        sbatofs=0
        self.SBATdata=[]
        self.SBATstreams=[]
        self.totalblocks=0
        for entry in self.properties[1:]:
            if 0<entry.size<self.minicutoff:
                entry.start = sbatofs
                self.SBATstreams.append(entry.data)
                blocks=(entry.size+(1<<self._miniblocksize)-1)>>self._miniblocksize
                self.SBATdata.extend(xrange(sbatofs+1,sbatofs+blocks))
                self.SBATdata.append(0xFFFFFFFE)
                sbatofs+=blocks
            else:
                self.totalblocks+=(entry.size+(1<<self._blocksize)-1)>>self._blocksize
        self.smallblocks = (sbatofs*(1<<self._miniblocksize)+(1<<self._blocksize)-1)>>self._blocksize
        self.totalblocks+= self.smallblocks
        self.sbatfatsize = (len(self.SBATdata)*4+(1<<self._blocksize)-1)>>self._blocksize
        self.totalblocks+= self.sbatfatsize
        self.properties[0].size = sbatofs*(1<<self._miniblocksize)
        
        
if False:
    for l in os.listdir('/home/user/workspace/Queck/mso/poi-3.8/test-data/spreadsheet/'):
        if l[-3:]=='xls':
            try:
                print l
                x=open('/home/user/workspace/Queck/mso/poi-3.8/test-data/spreadsheet/'+l)
                cfbr=CFBReader(x)
                cfbw=CFBWriter()
                for name,entry in cfbr.dirtree.iteritems():
                    data=entry.data[:entry.size]
                    cfbw.put(name, data)
            except Exception, e:
                print e
    cfbw.write('test.xls')
