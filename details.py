#
# -*- coding:utf-8 -*-
#出仓明细,其实这三个表应该可以继承自同一个类，这样会方便很多的。

class Details(object):
	
    HS_COL = "A" #hl code from db
    PN_COL = "B"
    NAME_COL = "C" #Chinese or English name from DB.. 注意居中
    AMOUNT_COL = "D" # the amount of export
    UNIT_PRICE_COL = "F" 
    TOTAL_PRICE_COL = "G"
	
    RECORD = [HS_COL, PN_COL, NAME_COL, AMOUNT_COL, UNIT_PRICE_COL,TOTAL_PRICE_COL]
    RECORD_NAME = ["hs","pn","name","amount","unit_price","total_price"]
	
    ROW_START = 4
    ROW_END = 24
    

	
    def __init__(self,sheet,**kwargs):
		self.sheet = sheet #complete here
		self.row = self.ROW_START
		
    #maybe more check here
    def insert_record(self, **kwargs):
	    cols = [item+str(self.row) for item in self.RECORD]
    	    if self.row >= self.ROW_END:
                assert 0, "Row overflow"
            self.row += 1    	
            for i in xrange(len(self.RECORD_NAME)):
                val = kwargs.get(self.RECORD_NAME[i],None)
                assert val != None, "Value of (%s) error when insert into invoice: "  % self.RECORD_NAME[i]  #directly end
                self.sheet[cols[i]] = val
            return 0

    def get_cell(self,pos):
        return self.sheet[pos].value

