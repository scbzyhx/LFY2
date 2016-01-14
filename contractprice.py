#
# -*- coding:utf-8 -*-
import logging

import utils
MAX_PNS = 20

NANO = u"NANO"
PC = u"微型电脑主机"

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
        self.logger = logging.getLogger(str(self.__class__))
        self.logger.setLevel(utils.LOG_LEVEL)
        tpi_pns = tpi_info.get_pns()
        ex_pns = ex_info.get_pns()
        
        if len(tpi_pns) != len(ex_pns) or len(tpi_pns) > MAX_PNS:
            self.ok = False
            return 
        for pn in tpi_pns:
            in_price = tpi_info.get_unit_price(pn)
            ex_price = ex_info.get_unit_price(pn)
            in_qty = tpi_info.get_qty(pn)
            ex_qty = ex_info.get_qty(pn)
        
            if ex_price is None or in_qty != ex_qty:
                self.ok = False
                self.data = {} #clear
                break
            assert in_qty == ex_qty #modify it here
            
            if in_price >= ex_price:
                self.data[pn] = {self.PRICE:in_price, self.QTY:in_qty}
            else: #ex_price > in_price
                self.data[pn] = {self.PRICE:round(ex_price*1.01, 2),self.QTY:ex_qty}  #TODO: precisious, two decimal
        
    def process_pac(self,pac_info,**kwargs):
        pns = self.data.keys()
        
        for pn in pns:
            pac_pcs = pac_info.get_pcs(pn)
            pac_nw = pac_info.get_nw(pn)
            pac_gw = pac_info.get_gw(pn)
            #print "pac_pcs=",pac_pcs, " pac_nw=", pac_nw," pac_gw", pac_gw, " qty=",self.data[pn][self.QTY]
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
            self.logger.info( "pn = %s, qty = %d, price = %f, nw = %f, gw = %f, model=%s, code=%s, desc=%s" \
            % (pn,self.data[pn][self.QTY], self.data[pn][self.PRICE], self.data[pn][self.NW], self.data[pn][self.GW], self.data[pn][self.MODEL], self.data[pn][self.CODE], self.data[pn][self.DESC]))
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
        nano, normal, pc = self.partition()
        return [item[0] for item in nano],[item[0] for item in normal], [item[0] for item in pc]
        
    
    def partition(self):
        nano = []
        normal = [] #each item is an tuple, (pn, code)
        pc = [] #each item is an tuple(pn,code)
        nano_price = []
        for pn in self.data.keys():
            desc = self.data[pn][self.DESC]
            desc = desc.upper()
            if desc.find(NANO) >=0:
                desc_list = desc.split(" ")
                desc_arr = desc_list[0:4]
                desc = " ".join(desc_arr)
                self.data[pn][self.DESC] = desc+"-CHN" #modify
                nano.append((pn,self.data[pn][self.CODE]))
                nano_price.append(self.data[pn][self.PRICE])
            elif desc.find(PC) >= 0:
                pc.append((pn,self.data[pn][self.CODE]))
            else:
                normal.append((pn,self.data[pn][self.CODE])) #to sorted
        normal.sort(key=lambda s: s[1]) #according to code
        pc.sort(key=lambda s: s[1])
        nano.sort(key=lambda s: s[1])
        for pn_pair in nano:
            self.data[pn_pair[0]][self.PRICE] = max(nano_price)
        
        return nano, normal, pc
    #nano is a list of pns, of whcih the desc is nano
    
    #PRICE = "price" , if not the same
    #QTY = "qty",
    #NW = "nw", sum?
    #GW = "gw", sum?
    #MODEL = "model, if not the same
    #CODE = "code", if not the same
    #DESC = "desc", must be the same
    def process_nano(self,nano):
        if len(nano) == 0:
            return {}
        record = {self.QTY:0, self.NW:0,self.GW:0}
        for pn in nano:
            if record.has_key(self.PRICE) and record[self.PRICE] != self.data[pn][self.PRICE]:
                self.logger.error(u"NANO 价格不相同")
                return None
            if record.has_key(self.CODE) and record[self.CODE] != self.data[pn][self.CODE]:
                self.logger.error(u" NANO 商品编码不相同")
                return None
                
            record[self.PRICE] = self.data[pn][self.PRICE] #如果  price 价格不一样怎么办？            
            record[self.QTY] += self.data[pn][self.QTY]
            
            record[self.DESC] = self.data[pn][self.DESC]
    def get_pns(self):
        return self.data.keys()
    def get_qty(self,pn):
        return self.data[pn][self.QTY]
    def get_price(self,pn):
        return self.data[pn][self.PRICE]
    def get_nw(self,pn):
        return self.data[pn][self.NW]
    def get_gw(self,pn):
        return self.data[pn][self.GW]
    def get_model(self,pn):
        return self.data[pn][self.MODEL]
    def get_code(self,pn):
        return self.data[pn][self.CODE]
    def get_desc(self,pn):
        return self.data[pn][self.DESC]
        
if __name__=="__main__":
    HelloWorld()
