#
# -*- coding:utf-8
import logging
import os
from datetime import datetime

from infodb import YInfoDB
from pacs import PACs
from export import Export
from tpis import TPIs
from output import OutputWorkbook
from contractprice import ContractPrice
from hawb import Account, HAWB

import utils
#TOP=u"C:\\Users\\yanghaixiang\\Desktop\\一键制单\\"
TOP=u".\\workspace\\"
INPUT=TOP+u"INPUT\\"
OUTPUT=TOP+u"OUTPUT\\"

INFO_DB = INPUT+u"DB\\规范申报数据库.xlsx"
TPI_DIR = INPUT+u"TPI\\"
EXPORT = INPUT#+u"出口单价与数量.xlsx"
PAC_DIR = INPUT+u"PAC\\"

TEMPLATE_FILE = TOP+u"模板.xlsx"
CONTRACT_ID = None
DEST_FILE = OUTPUT#+u"sample.xlsx"
ACCOUNT_FILE = OUTPUT+u"记账.xlsx"
HAWB_FILE = OUTPUT+u"HAWB.xlsx"

##initialize logging
logging.basicConfig()

LOG = logging.getLogger(__name__)
LOG.setLevel(utils.LOG_LEVEL)
#process nano list
#ct is contract sheet
#dt is detail sheet
#it is invoice sheet
#nano is the P/N list of nano,
#cprice is the contractPrice instance        
def process_nano(ct,dt,it,nano,cprice,info):
    total_qty = 0
    nw = 0
    total_amount = 0
    if len(nano) == 0:
        return
    for pn in nano:
        qty = cprice.get_qty(pn)
        
        total_qty += qty
        desc = cprice.get_desc(pn)
        nw = cprice.get_nw(pn)
        code = cprice.get_code(pn)
        model = cprice.get_model(pn)
        price = cprice.get_price(pn)
        
        amount = round(price*qty,2)
        total_amount += amount
        invoice_record = {
            "pn":pn, #useless
            "model":model,
            "description":info.get_desc(pn),
            "qty":qty,
            "nw":nw,
            "unit_price":price,
            "amount":amount
        }
        it.insert_record(**invoice_record)
    d_record={
            "hs":code,
            "pn":pn, #useless
            "name":desc,
            "amount":qty,
            "unit_price":price,
            "total_price":total_amount
    }
    
    dt.insert_record(**d_record)
    ct.insert_record(total_qty,desc,price,total_amount)

##process normal list or pc list
def process_other(ct,dt,it,pns,c_price):
    #invoice:RECORD_NAME = ["pn","model","description","qty","nw","unit_price","amount"]
    #details: RECORD_NAME = ["hs","pn","name","amount","unit_price","total_price"]
    for pn in pns:
        qty = c_price.get_qty(pn)
        desc = c_price.get_desc(pn)
        price = c_price.get_price(pn)
        nw = c_price.get_nw(pn) #净重
        #gw = c_price.get_gw(pn) #毛重
        model = c_price.get_model(pn)  #model 型号
        code = c_price.get_code(pn) #商品编码
        
        amount = round(price*qty,2)
        invoice_record = {
            "pn":pn,
            "model":model,
            "description":desc,
            "qty":qty,
            "nw":nw,
            "unit_price":price,
            "amount":amount
        }
        d_record={
            "hs":code,
            "pn":pn,
            "name":desc,
            "amount":qty,
            "unit_price":price,
            "total_price":amount
        }
        it.insert_record(**invoice_record)
        #print invoice_record
        dt.insert_record(**d_record)
        ct.insert_record(qty,desc,price,qty*price)
    return 0
    
def process_others(ct,dt,it,contract_id,pacs,cprice):
    date = datetime.now().strftime("%Y/%m/%d")
    gw = 0
    total_qty = 0
    total_nw = 0
    total_amount = 0
    
    pns = cprice.get_pns()
    for pn in pns:
        gw += cprice.get_gw(pn)
        total_qty += cprice.get_qty(pn)
        total_nw += cprice.get_nw(pn)
        total_amount += cprice.get_price(pn)*cprice.get_qty(pn)
    
    
    
    
    ct.insert_contract_id(contract_id)
    ct.insert_contract_date(date)
    ct.set_total_amount(total_amount)
    ct.set_total_qty(total_qty)
    
    it.insert_contract_id(contract_id)
    it.insert_contract_date(date)    
    it.insert_plts(pacs.get_plts())
    it.insert_gw(gw)
    it.set_total_qty(total_qty)
    it.set_total_amount(total_amount)
    it.set_total_nw(total_nw)
    
    dt.set_total_qty(total_qty)
    dt.set_total_amount(total_amount)
    
def save_account(tpis, exports, contracts):
    accounts = Account(ACCOUNT_FILE)
    
    pns = contracts.get_pns()
    for pn in pns:
        iprice = tpis.get_unit_price(pn)
        oprice = exports.get_unit_price(pn)
        cprice = contracts.get_price(pn)
        
        accounts.insert_record(CONTRACT_ID,pn,iprice,oprice,cprice)
    accounts.save()#save files

def save_hawb_total(tpis, contracts):
    all_tpi =  tpis.get_all_tpi()
    hawb = HAWB(HAWB_FILE)
    for tpi in all_tpi:
        amount = 0
        pns = tpi.get_pns()
        for pn in pns:
            amount += tpi.get_qty(pn)*contracts.get_price(pn)
        hawb.insert_record(CONTRACT_ID,tpi.get_hawb(),amount)
    hawb.save()
    
def main():
    info_db = YInfoDB(INFO_DB)
    tpis_info = TPIs(TPI_DIR)    
    exports = Export(EXPORT)
    pacs_info = PACs(PAC_DIR)
    
    c_price = ContractPrice(tpis_info,exports)
    if not c_price.is_ok():
        LOG.error("Error")
        return 
        
    #pacs_info.show_pacs()
    
    if c_price.process_pac(pacs_info) == False:
        LOG.error("Error When process PACs")
        return
    if c_price.process_infodb(info_db) == False:
        LOG.error("Error when process info db")
        return 
    c_price.show_contract_price()
    
    outbook = OutputWorkbook(TEMPLATE_FILE)
    nanos,normal_pns, pc_pns = c_price.get_data()
    
    ct = outbook.get_contract_sheet()
    dt = outbook.get_detail_sheet()
    it = outbook.get_invoice_sheet()
    
    LOG.info("nano=%s",nanos)
    LOG.info("normal=%s",normal_pns)
    LOG.info("pc=%s",pc_pns)
    #first process nano
    process_nano(ct,dt,it,nanos,c_price,info_db)
    #second normal
    process_other(ct,dt,it,normal_pns,c_price)
        
    #third, 微型电脑主机
    process_other(ct,dt,it,pc_pns,c_price)
    
    process_others(ct,dt,it,CONTRACT_ID,pacs_info,c_price)

    
    outbook.save(DEST_FILE)
    
    ##now accouting
    save_account(tpis_info,exports,c_price)
    save_hawb_total(tpis_info,c_price)
    print u"处理成功!!"
    utils.pause()
def find_contract_id():
    files = os.listdir(INPUT)
    for fl in files:
        if os.path.isfile(INPUT+fl) == False:
            continue
        fls = fl.split(".")
        if fls[-1] == "xlsx":
            global DEST_FILE
            global CONTRACT_ID
            global EXPORT
            CONTRACT_ID = fls[0]
            DEST_FILE += CONTRACT_ID+u"import.xlsx"
            EXPORT += CONTRACT_ID + u".xlsx"
            LOG.info(u"合同号为:%s",CONTRACT_ID)
            break
            
if __name__=="__main__":
    print u"版权所有"
    find_contract_id()
    main()
