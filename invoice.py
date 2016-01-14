#
# -*- coding:utf-8 -*-
import logging
import sys
import utils

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
    TOTAL_QTY_POS = "G28"
    TOTAL_NW_POS = "H28"
    TOTAL_AMOUNT_POS = "J28"
    
    RECORD = [PN_COL,MODEL_COL,DESC_COL,QTY_COL,NW_COL,PRICE_COL,AMOUNT_COL]
    RECORD_NAME = ["pn","model","description","qty","nw","unit_price","amount"]
    
    
    
    def __init__(self,sheet,**kwargs):
        self.sheet = sheet #complete here
        self.row = self.ROW_START
        self.logger = logging.getLogger(str(self.__class__))
        self.logger.setLevel(utils.LOG_LEVEL)
        
    def insert_record(self,**kwargs):
    	cols = [item+str(self.row) for item in Invoice.RECORD]
    	if self.row >= self.ROW_END:
    	#	print "PN overflow"
            self.logger.error(u"列数超过了 20 个")
            #utils.show_abort()
            sys.exit(-1)
    		
    	self.row += 1    	
    	for i in xrange(len(self.RECORD_NAME)):
            val = kwargs.get(self.RECORD_NAME[i],None)
            if val is None:
                self.logger.error("Value of (%s) error when insert into invoice: ",self.RECORD_NAME[i])
                sys.exit(-1)
    		#assert val != None, #directly end
            #print cols[i],"   ",val
    	    self.sheet[cols[i]] = val
    	return (0,"ok")
    def remove_empty_rows(self): #copy
        #test 
        #self.row = 22
        last_row = self.sheet.get_highest_row() 
        #last_col = 8#self.sheet.get_highest_column() + 1
        #print last_col
        for i in xrange(self.row,self.ROW_END):
            self.sheet.row_dimensions[i].hidden = True   
             		    	
    def insert_contract_id(self,c_id):
    	self.sheet[self.CONTRACT_ID_POS] = c_id
    def insert_contract_date(self,date):
    	self.sheet[self.CONTRACT_DATE_POS] = date
    def insert_plts(self,plts):
    	self.sheet[self.PLTS_POS] = plts
    def insert_gw(self,gw):
    	self.sheet[self.GW_POS] = gw
    def set_total_qty(self,total_qty):
        self.sheet[self.TOTAL_QTY_POS] = total_qty
    def set_total_nw(self,total_nw):
        self.sheet[self.TOTAL_NW_POS] = total_nw
    def set_total_amount(self,total_amount):
        self.sheet[self.TOTAL_AMOUNT_POS] = total_amount
