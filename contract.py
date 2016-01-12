#
# -*- coding:utf-8 -*-
from openpyxl.styles import Border,Side,Alignment, Protection, Font

class Contract(object):
    QTY_COL = "A"
    DESC_COL = "B"
    PRICE_COL = "F"
    AMOUNT_COL = "G"
    ROW_START = 17
    ROW_END = 37

    CONTRACT_ID_POS = "F4"
    CONTRACT_DATE_POS = "F8"
    def __init__(self,sheet,**kwargs):
    	super(Contract,self).__init__()
        self.sheet = sheet #complete here
        self.row = 17

    def insert_record(self,qty,desc,price,amount):
        if self.row >= self.ROW_END:
    		return -1
    		
        self.sheet[self.QTY_COL+str(self.row)] = qty
        self.sheet[self.DESC_COL+str(self.row)] = desc 
        self.sheet[self.PRICE_COL+str(self.row)] = price
        self.sheet[self.AMOUNT_COL+str(self.row)] = amount

        #OK, future
        desc_cell = self.sheet[self.DESC_COL+str(self.row)]
        desc_cell.alignment = Alignment(horizontal='center')
        
    	
        return 0

    def insert_contract_id(self,c_id):
    	self.sheet[self.CONTRACT_ID_POS] = c_id
        

    def insert_contract_date(self,date):
        self.sheet[self.CONTRACT_DATE_POS] = date

    def close(self):
        "save"
        pass
        
    def get_cell(self,pos=QTY_COL+str(ROW_START)):
    	return self.sheet[pos].value#.encode('ascii','ignore')
    def merged_cells(self):
        self.sheet.merge_cells('B17:D17')



if __name__ == "__main__":
    print u"人是人，鬼是鬼"
    from openpyxl import load_workbook
    try:
    	wb = load_workbook(filename=u"..\\空白模板123.xlsx")
    	contract = Contract(wb.get_sheet_by_name(u"合同"))
    	contract.show_cell("A16")
    	print contract
    except IOError :
    	print "err"
