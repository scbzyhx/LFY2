#
# -*- coding:utf-8
import logging
import os
from openpyxl import load_workbook
from rworkbook import YRWorkbook
MAX_TIMES = 5
#logging.basicConfig()
#LOG = logging.getLogger(__name__)
#LOG.setLevel(logging.DEBUG)
import utils

class TPI(object):
    PN_COL = "A"
    DESC_COL = "B"
    QTY_COL = "J"
    UNIT_PRICE_COL = "K"
    EXTENDED_COL = "L"
    
    HAWB_POS = "A9"
    
    START_ROW = 13 #why start from 13
    
    END_STR = "Total"
    
    def __init__(self,filename, hwab,**kwargs):
        self.filename = filename
        self.hwab = hwab
        self.data = {} #pn: [mxprice,sum of pty] #does we need extend
        self.logger = logging.getLogger("TPI")
        self.logger.setLevel(utils.LOG_LEVEL)
        #print self.hwab
        self.hawb = 0
        self._process_tpi()
    def _get_hawb(self,s):
        #print s
        self.hawb = s.split("#")[-1].strip()
        #print self.hawb
        assert self.hawb != ""
        
    def get_hawb(self):
        return self.hawb
        
    def get_pns(self):
        return self.data.keys()
    def get_qty(self,pn):
        if self.data.has_key(pn):
            return self.data[pn][1]
    
    def _process_tpi(self):
        try:
            wb = YRWorkbook(self.filename)
        except IOError :
            utils.ABT(u"处理文件 %s 出错" % self.filename)
        names = wb.get_sheet_names()
        
        for name in names:
            times = 0
            #ind = self.START_ROW
            wb.pin_sheet_by_name(name)
            nrows = wb.get_nrows()
            self.logger.debug(name)
            self._get_hawb(wb.get_cell_value(self.HAWB_POS))
            for ind in xrange(self.START_ROW, nrows):
                pn = wb.get_cell_value(self.PN_COL+str(ind))
                qty = wb.get_cell_value(self.QTY_COL+str(ind))
                uprice =wb.get_cell_value(self.UNIT_PRICE_COL+str(ind))
                extend = wb.get_cell_value(self.EXTENDED_COL+str(ind))
                
                extend_type = wb.get_cell_type(self.EXTENDED_COL+str(ind))
                if extend == self.END_STR:
                    break
                if extend  ==  "":
                    times += 1
                    continue
                if times >=  MAX_TIMES:
                    break
                    
                #ok, then clear timer
                times = 0
                if not self.data.has_key(pn):
                    self.data[pn] = [0,0]
                if uprice[0] != "$":
                    self.logger.debug("pn=%s, unit_price=%s, pty=%s",pn, uprice, qty)
                    self.logger.warn("Unit_Price Error, unit price shuold start with $")
                    continue
                
                if int(qty) <= 0:
                    utils.ABT(u"ERROR: TPI 中, 这条记录（pn=%s,qty=%s,unit_price=%s） 有错误" %(pn,qty,uprice))
                    
                    
                self.data[pn][0] = max(self.data[pn][0],float(uprice[1:]))
                self.data[pn][1] += int(qty)
                
                self.logger.debug("pn=%s, unit_price=%f, pty=%d",pn, float(uprice[1:]), int(qty))
    def print_sumary(self):
        for k in self.data.keys():
            self.logger.debug("Statistics:hwab = %s, pn=%s, unit_price=%f, pty=%d",self.hwab, k, self.data[k][0], self.data[k][1])


class TPIs(object):
    def __init__(self,dir, **kwargs):
        self.dir = dir
        self.logger = logging.getLogger("TPIs")
        self.logger.setLevel(utils.LOG_LEVEL)
        self.tpi_list = []
        self.db = {}
        self._process(self.dir)
        self._sumary()
    def _process(self,dir):
        files = os.listdir(dir)
        #print files
        #print dir
        for fl in files:
            #self.logger.info(fl)
            if os.path.isfile(dir+fl) == False:
                #self.logger.info("非文件: %s" % (dir+fl))
                continue
            #filename = 
            #print filename
            #print "abcd"
            if fl.split('.')[-1] != "xlsx" and fl.split('.')[-1] != "xls":
                #print fl
                self.logger.warn(u"TPI 目录中包含非 Excel 文件: %s    [已忽略]" % (fl))
                continue
            #print fl            
            self.tpi_list.append(TPI(dir+fl,fl.split(".")[0].split("_")[1]))
            
        if len(self.tpi_list) == 0:
            utils.ABT(u"ERROR: TPI 文件缺失！")
            
    def _sumary(self):
        for tpi in self.tpi_list:
            tpi.print_sumary()
            keys = tpi.data.keys()
           # print tpi.data
            for k in keys:
                if not self.db.has_key(k):
                    ##OR DEEP COPY HERE
                    self.db[k] = [0,0]
                    self.db[k][0] = tpi.data[k][0]
                    self.db[k][1] = tpi.data[k][1]
                else:
                    self.db[k][1] += tpi.data[k][1]
                    self.db[k][0] = max(tpi.data[k][0], self.db[k][0])
    def get_pns(self):
        return self.db.keys()
    
    def get_unit_price(self,pn):
        if self.db.has_key(pn):
            return self.db[pn][0]
        return None
        
    def get_qty(self,pn):
        if self.db.has_key(pn):
            return self.db[pn][1]
        return None
        
    def get_all_tpi(self):
        return self.tpi_list
                
if __name__=="__main__":
    #tpi =  TPI("C:\\Users\\yanghaixiang\\Documents\\GitHub\\TPI_VC70160106A07.xls","123")
    tpis = TPIs(u"C:\\Users\\yanghaixiang\\Documents\\GitHub\\LFY2\\workspace\\IN\\T\\")
    pns  = tpis.get_pns()
    for tpi in tpis.get_all_tpi():
        print tpi.data
    for pn in pns:
        print pn,"   ",tpis.get_unit_price(pn),"  ", tpis.get_qty(pn)
        
