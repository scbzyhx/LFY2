#
# -*- coding:utf-8
import logging
import os

from openpyxl import load_workbook

from contract import Contract
from invoice import Invoice
from details import Details

class OutputWorkbook:
    CONTRACT_SHEET = u"合同"
    INVOICE_SHEET = u"箱单发票"
    DETAILS_SHEET = u"出仓明细"
    def __init__(self,template, **kwargs):
        self.template = template
        #self.dest = dest
        
        self.wb = load_workbook(template)
        
        self.contract = Contract(self.wb.get_sheet_by_name(self.CONTRACT_SHEET))
        self.invoice = Invoice(self.wb.get_sheet_by_name(self.INVOICE_SHEET))
        self.details = Details(self.wb.get_sheet_by_name(self.DETAILS_SHEET))
        
    def get_contract_sheet(self):
        return contract
    def get_invoice_sheet(self):
        return self.invoice
    def get_detail_sheet(self):
        return self.details
        
    def save(self,dest=None):
        if dest is None:
            dest = self.template
        #save here

if __name__=="__main__":
    pass
