#
# -*- coding:utf-8

from openpyxl import load_workbook

from contract import Contract
from invoice import Invoice
from details import Details

class Clearance(object):
    CONTRACT = u"合同"
    INVOICE = u"箱单发票"
    DETAILS = u"出仓明细"
	#may raise IOError exception
    def __init__(self, filename, **kwargs):
        self.filename = filename
        self.wb = load_workbook(filename)
		
        self.contract = Contract(self.wb.get_sheet_by_name(self.CONTRACT))
        self.invoice = Invoice(self.wb.get_sheet_by_name(self.INVOICE))
        self.details = Details(self.wb.get_sheet_by_name(self.DETAILS))
		
		
    def get_contract_table(self):
		return self.contract
		
    def get_invoice_table(self):
		return self.invoice
    def get_details_table(self):
		return self.details
		
    def save(self):
		self.wb.save(filename=self.filename)
    def __del__(self):
		pass
		

if __name__ == "__main__":
    clr = Clearance(u"..\\空白模板.xlsx")
    contract = clr.get_contract_table()
    invoice = clr.get_invoice_table()
    print contract.get_cell()
    
    contract.insert_record(1,"aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",3,4)
    contract.merged_cells()

    detail = clr.get_details_table()
    records = {
        "hs":"1235",
        "pn":"6789",
        "name":"hello world",
        "amount":987,
        "unit_price":1230,
        "total_price":1234
        
    }
    detail.insert_record(**records)
    #invoice
    records = {
    "pn":"123",
    "model":"09876",
    "description":"0987",
    "qty":01234,
    "nw":0123,
    "unit_price":9807,
    "amount":90008
    }
    invoice.insert_record(**records)
    
    contract.insert_contract_id("0x1234567890")
    contract.insert_contract_date("2016/01/07")
    invoice.insert_contract_id("0x1234567890")
    invoice.insert_contract_date("2016/01/07")
    invoice.insert_plts(12354)
    invoice.insert_gw(1234)
    clr.save()
