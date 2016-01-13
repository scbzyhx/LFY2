#
# -*- coding:utf-8 -*-
#出仓明细,其实这三个表应该可以继承自同一个类，这样会方便很多的。
import utils
import logging
import sys

class Details(object):
	
    HS_COL = "A" #hl code from db
    #PN_COL = "B"
    NAME_COL = "B" #Chinese or English name from DB.. 注意居中
    AMOUNT_COL = "C" # the amount of export
    UNIT_PRICE_COL = "E" 
    TOTAL_PRICE_COL = "F"
	
    RECORD = [HS_COL,  NAME_COL, AMOUNT_COL, UNIT_PRICE_COL,TOTAL_PRICE_COL] #remove PN_COL,
    RECORD_NAME = ["hs","name","amount","unit_price","total_price"] #remove ,"pn"
	
    ROW_START = 4
    ROW_END = 24
    
    TOTAL_QTY_POS = "C24"
    TOTAL_AMOUNT_POS = "F24"
    
	
    def __init__(self,sheet,**kwargs):
        self.sheet = sheet #complete here
        self.row = self.ROW_START
        self.logger = logging.getLogger(str(self.__class__))
		
    #maybe more check here
    def insert_record(self, **kwargs):
	    cols = [item+str(self.row) for item in self.RECORD]
    	    if self.row >= self.ROW_END:
                
                self.logger.error(u"列数超过了 20 个")
                utils.show_abort()
                sys.exit(-1)
                
            self.row += 1    	
            for i in xrange(len(self.RECORD_NAME)):
                val = kwargs.get(self.RECORD_NAME[i],None)
                assert val != None, "Value of (%s) error when insert into invoice: "  % self.RECORD_NAME[i]  #directly end
                self.sheet[cols[i]] = val
            return 0

    def get_cell(self,pos):
        return self.sheet[pos].value
    def set_total_qty(self,total_qty):
        self.sheet[self.TOTAL_QTY_POS] = total_qty
    def set_total_amount(self,total_amount):
        self.sheet[self.TOTAL_AMOUNT_POS] = total_amount
    def remove_empty_rows(self): #copy
        #test 
        #self.row = 22
        last_row = self.sheet.get_highest_row() 
        #last_col = 8#self.sheet.get_highest_column() + 1
        #print last_col
        for i in xrange(self.row,self.ROW_END):
            self.sheet.row_dimensions[i].hidden = True
    def fix_borders(self):
        border = self.sheet["B3"].border
        self.sheet["C23"].border = border
        self.sheet["B24"].border = border
        self.sheet["C24"].border = border

