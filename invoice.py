#
# -*- coding:utf-8 -*-

class Invoice(object):
    PN_COL = "B"
    MODEL_COL = "C"
    DESC_COL = "D"
    QTY_COL = "G"
    NW_COL = "H"
    PRICE_COL = "I"
    AMOUNT_COL = "J"
    ROW_START = 8
    ROW_END = 28
    
    CONTRACT_ID_POS = "J2"
    CONTRACT_DATE_POS = "J3"
    
    PLTS_POS = "G31" #托数
    GW_POS = "H31"   #毛重
    
    RECORD = [PN_COL,MODEL_COL,DESC_COL,QTY_COL,NW_COL,PRICE_COL,AMOUNT_COL]
    RECORD_NAME = ["pn","model","description","qty","nw","unit_price","amount"]
    
    
    
    def __init__(self,sheet,**kwargs):
        self.sheet = sheet #complete here
        self.row = self.ROW_START
        
    def insert_record(self,**kwargs):
    	cols = [item+str(self.row) for item in Invoice.RECORD]
    	if self.row >= self.ROW_END:
    	#	print "PN overflow"
    		return (-1,"overflow")
    		
    	self.row += 1    	
    	for i in xrange(len(self.RECORD_NAME)):
    		val = kwargs.get(self.RECORD_NAME[i],None)
    		assert val != None, "Value of (%s) error when insert into invoice: "  % self.RECORD_NAME[i]#directly end
    		self.sheet[cols[i]] = val
    	return (0,"ok")
    		    	
    def insert_contract_id(self,c_id):
    	self.sheet[self.CONTRACT_ID_POS] = c_id
    def insert_contract_date(self,date):
    	self.sheet[self.CONTRACT_DATE_POS] = date
    def insert_plts(self,plts):
    	self.sheet[self.PLTS_POS] = plts
    def insert_gw(self,gw):
    	self.sheet[self.GW_POS] = gw
