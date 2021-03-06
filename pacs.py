#
# -*- coding:utf-8
import logging
import os

from rworkbook import YRWorkbook

import utils

class PAC:
    
    PN_COL = "B"
    PCS_COL = "F"
    NW_COL = "G"
    GW_COL = "H"
    END_TAG_COL = "E"
    PLT_COL = "A"
    
    START_ROW = 13
    END_STR = "total"
    
    PCS = "pcs"
    NW = "nw"
    GW = "gw"
    

    def __init__(self,filename,hwab,**kwargs):
        #super(PAC,self).__init()
        self.logger = logging.getLogger("PAC")
        self.logger.setLevel(utils.LOG_LEVEL)
        self.filename = filename
        self.hwab = hwab
        self.data = {}
        self.plt = 0
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
                if pcs <= 0 or nw <= 0 or gw <= 0:
                    utils.ABT(u"ERROR: PAC 中，这条记录(pn=%s,pcs=%d,nw=%f,gw=%f)存在错误" %(pn,pcs,nw,gw))

                
                if not end_tag.find(self.END_STR):
                    self._get_plt(wb,self.PLT_COL+str(ind))
                    break
                self.logger.debug("%s, %d, %f, %f, %s", pn, pcs, nw, gw,end_tag)
                
                if not self.data.has_key(pn):
                    self.data[pn] = {self.PCS: 0, self.NW:0.0, self.GW:0.0}
                self.data[pn][self.PCS] += pcs
                self.data[pn][self.NW] += nw
                self.data[pn][self.GW] += gw

    #wb is an instance of YRWorkbook
    def _get_plt(self,wb, pos):
        #process cell
        plt_str = wb.get_cell_value(pos)
        plt_str_list = plt_str.split(" ")
        if plt_str_list[0].lower() == "total:" and plt_str_list[2].lower().find("plt") >= 0:
            self.plt += int(plt_str_list[1])
        else:
            self.plt = -1;
        
    def get_plt(self):
        return self.plt
        
    def printf(self):
        for k,v in self.data.items():
            print self.data[k][self.NW]
            
            
class PACs(object):
    def __init__(self,dir,**kwargs):
        self.dir = dir
        self.logger = logging.getLogger("PACs")
        self.logger.setLevel(logging.DEBUG)
        self.pac_list = []
        self.db = {}
        self.plts = 0
        self._process(dir)
        self._sumary()
        
        
    def _process(self,dir):
        files = os.listdir(dir)
        #print files
        for fl in files:
            #print fl
            if os.path.isfile(dir+fl) == False:
                #self.logger.info("illegal file %s" % (dir+fl))
                continue
            #filename = 
            #print filename
            if fl.split('.')[-1] != "xlsx" and fl.split('.')[-1] != "xls":
                self.logger.warn(u"PAC 目录中包含非 Excel 文件: %s    [已忽略]" % (fl))
                continue
            #print fl
            
            self.pac_list.append(PAC(dir+fl,fl.split(".")[0].split("_")[1]))
        if len(self.pac_list) == 0:
            utils.ABT(u"ERROR: PAC 文件缺失")
    def _sumary(self):
        for pac  in self.pac_list:
            keys = pac.data.keys()
            
            if self.plts >= 0 and pac.get_plt() > 0:
                self.plts += pac.get_plt()
            
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
    def get_plts(self):
        return self.plts#total plts
    
            
if __name__=="__main__":
   #pac = PAC(u"C:\\Users\\yanghaixiang\\Desktop\\一键制单\\PAC\\Pac_CPVC64150720A05.xls","123")
   pacs = PACs(u"C:\\Users\\yanghaixiang\\Desktop\\一键制单\\PAC\\")
   #pac.printf()
   print pacs.get_plts()
       
   print pacs.get_pcs("Z0RJ000ET")
   print pacs.get_nw("Z0RJ000ET")
   print pacs.get_gw("Z0RJ000ET")
   
   print pacs.get_pcs("Z0RJ000E")