#
# -*- coding:utf-8
import logging
import os

from infodb import YInfoDB
from pacs import PACs
from export import Export
from tpis import TPIs


TOP=u"C:\\Users\\yanghaixiang\\Desktop\\一键制单\\"

INFO_DB = TOP+u"规范申报数据库.xlsx"
TPI_DIR = TOP+u"TPI\\"
EXPORT = TOP+u"出口单价与数量.xlsx"
PAC_DIR = TOP+u"PAC\\"
##initialize logging
logging.basicConfig()

#
#tpi_info is an instance of TPIs while ex_info is an instance of  Export
class ContractPrice:
    PRICE = "price"
    QTY = "qty"
    NW = "nw"
    GW = "gw"
    MODEL = "model"
    CODE = "code"
    DESC = "desc"
    def __init__(self,tpi_info, ex_info,**kwargs):
        self.data = {} #None
        self.ok = True
        
        tpi_pns = tpi_info.get_pns()
        ex_pns = ex_info.get_pns()
        
        if len(tpi_pns) != len(ex_pns):
            return None
        for pn in tpi_pns:
            in_price = tpi_info.get_unit_price(pn)
            ex_price = ex_info.get_unit_price(pn)
            in_qty = tpi_info.get_qty(pn)
            ex_qty = ex_info.get_qty(pn)
        
            if ex_price is None or in_qty != ex_qty:
                self.ok = False
                self.data = {} #clear
                break
            assert in_qty == ex_qty
            
            if in_price >= ex_price:
                self.data[pn] = {self.PRICE:in_price, self.QTY:in_qty}
            else: #ex_price > in_price
                self.data[pn] = {self.PRICE:ex_price*1.01,self.QTY:ex_qty}  #TODO: precious.##############
        
    def process_pac(self,pac_info,**kwargs):
        pns = self.data.keys()
        
        for pn in pns:
            pac_pcs = pac_info.get_pcs(pn)
            pac_nw = pac_info.get_nw(pn)
            pac_gw = pac_info.get_gw(pn)
            print "pac_pcs=",pac_pcs, " pac_nw=", pac_nw," pac_gw", pac_gw, " qty=",self.data[pn][self.QTY]
            if pac_pcs != self.data[pn][self.QTY]:
                self.ok = False
                return False
            self.data[pn][self.NW] = pac_nw
            self.data[pn][self.GW] = pac_gw
        return True
    def is_ok(self):
        return self.ok
    def show_contract_price(self):
        for pn in self.data.keys():
            print "pn = %s, qty = %d, price = %f, nw = %f, gw = %f, model=%s, code=%s, desc=%s" \
            % (pn,self.data[pn][self.QTY], self.data[pn][self.PRICE], self.data[pn][self.NW], self.data[pn][self.GW], self.data[pn][self.MODEL], self.data[pn][self.CODE], self.data[pn][self.DESC])
    def process_infodb(self,infodb,**kwargs):
        for pn in self.data.keys():
            model = infodb.get_model(pn)
            code = infodb.get_code(pn)
            desc = infodb.get_desc(pn)
            
            if model is None or code is None or desc is None:
                self.ok = False
                return False
            self.data[pn][self.MODEL] = model
            self.data[pn][self.CODE] = code
            self.data[pn][self.DESC] = desc
        return True
    #return as tuples
    def get_data(self):
        return None
def main():
    info_db = YInfoDB(INFO_DB)
    tpis_info = TPIs(TPI_DIR)    
    exports = Export(EXPORT)
    pacs_info = PACs(PAC_DIR)
    
    c_price = ContractPrice(tpis_info,exports)
    if not c_price.is_ok():
        print "Error"
        return 
    #pacs_info.show_pacs()
    
    if c_price.process_pac(pacs_info) == False:
        print "Error When process pacs"
        return
    if c_price.process_infodb(info_db) == False:
        print "Error when process info db"
        return 
    c_price.show_contract_price()
    
    
    print "success"
    

if __name__=="__main__":
    main()
