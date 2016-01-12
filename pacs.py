#
# -*- coding:utf-8
import logging
import os

from rworkbook import YRWorkbook


class PAC:
    
    PN_COL = "B"
    PCS_COL = "F"
    NW_COL = "G"
    GW_COL = "H"
    END_TAG_COL = "E"
    
    START_ROW = 13
    END_STR = "total"
    
    PCS = "pcs"
    NW = "nw"
    GW = "gw"

    def __init__(self,filename,hwab,**kwargs):
        #super(PAC,self).__init()
        self.logger = logging.getLogger("PAC")
        self.logger.setLevel(logging.DEBUG)
        self.filename = filename
        self.hwab = hwab
        self.data = {}
        self._process_pac()
        
    def _process_pac(self):
        wb = YRWorkbook(self.filename)
        names = wb.get_sheet_names()
        
        for name in names:
            wb.pin_sheet_by_name(name)
            self.logger.debug(name)
            nrows = wb.get_nrows()
            
            for ind in xrange(self.START_ROW,nrows):
                pn = wb.get_cell_value(self.PN_COL + str(ind))
                pcs = wb.get_cell_value(self.PCS_COL + str(ind))
                nw = wb.get_cell_value(self.NW_COL + str(ind))
                gw = wb.get_cell_value(self.GW_COL + str(ind))
                end_tag = wb.get_cell_value(self.END_TAG_COL + str(ind)).lower()
                
                pcs = int(pcs)
                nw = float(nw)
                gw = float(gw)
                #print type(pcs)
                
                assert pcs != 0
                assert nw != 0
                assert gw != 0
                
                if not end_tag.find(self.END_STR):
                    break
                self.logger.debug("%s, %d, %f, %f, %s", pn, pcs, nw, gw,end_tag)
                
                if not self.data.has_key(pn):
                    self.data[pn] = {self.PCS: 0, self.NW:0.0, self.GW:0.0}
                self.data[pn][self.PCS] += pcs
                self.data[pn][self.NW] += nw
                self.data[pn][self.GW] += gw
    def printf(self):
        for k,v in self.data.items():
            print self.data[k][self.NW]
            
            
class PACs(object):
    def __init__(self,dir,**kwargs):
        self.dir = dir
        self.logger = logging.getLogger(str(self.__class__))
        self.logger.setLevel(logging.DEBUG)
        self.pac_list = []
        self.db = {}
        self._process(dir)
        self._sumary()
        
    def _process(self,dir):
        files = os.listdir(dir)
        #print files
        for fl in files:
            #print fl
            if os.path.isfile(dir+fl) == False:
                self.logger.info("illegal file %s" % (dir+fl))
                continue
            #filename = 
            #print filename
            if fl.split('.')[-1] != "xlsx" and fl.split('.')[-1] != "xls":
                self.logger.info("illegal file %s" % (dir+fl))
                continue
            #print fl
            
            self.pac_list.append(PAC(dir+fl,fl.split(".")[0].split("_")[1]))

    def _sumary(self):
        for pac  in self.pac_list:
            keys = pac.data.keys()
            for k in keys:
                if not self.db.has_key(k):
                    self.db[k] = {PAC.PCS: 0, PAC.NW:0.0, PAC.GW:0.0}
                self.db[k][pac.PCS] += pac.data[k][pac.PCS]
                self.db[k][pac.NW] += pac.data[k][pac.NW]
                self.db[k][pac.GW] += pac.data[k][pac.GW]
    def show_pacs(self):
        for k in self.db.keys():
            print k," ",self.db[k]["pcs"]," " ,self.db[k]["nw"], "  ",self.db[k]["gw"]
    def get_pcs(self,pn):
        if self.db.has_key(pn):
            return self.db[pn][PAC.PCS]
    def get_nw(self,pn):
        if self.db.has_key(pn):
            return self.db[pn][PAC.NW]
    def get_gw(self,pn):
        if self.db.has_key(pn):
            return self.db[pn][PAC.GW]
            
if __name__=="__main__":
   #pac = PAC(u"C:\\Users\\yanghaixiang\\Desktop\\一键制单\\PAC\\Pac_CPVC64150720A05.xls","123")
   pacs = PACs(u"C:\\Users\\yanghaixiang\\Desktop\\一键制单\\PAC\\")
   #pac.printf()
   
       
   print pacs.get_pcs("Z0RJ000ET")
   print pacs.get_nw("Z0RJ000ET")
   print pacs.get_gw("Z0RJ000ET")
   
   print pacs.get_pcs("Z0RJ000E")