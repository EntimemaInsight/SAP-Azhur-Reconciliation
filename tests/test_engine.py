import datetime as dt
import unittest
from reconciliation.engine import canonical_amount,match_documents,match_suppliers,normalize_invoice,normalize_supplier,reconcile_suppliers

def rec(system,sid,name,amount,invoice="INV-1",date=dt.date(2026,1,1),currency="EUR"):
    return {"source_system":system,"supplier_source_id":sid,"supplier_name_raw":name,"supplier_name_normalized":normalize_supplier(name),"signed_amount":amount,"invoice_normalized":normalize_invoice(invoice),"invoice_raw":invoice,"document_date":date,"currency":currency,"document_number":"1","reference":invoice,"assignment":"","due_date":date}

class EngineTests(unittest.TestCase):
    def test_supplier_normalization(self):self.assertEqual(normalize_supplier('  „Алфа-Бета“ ЕООД '),"АЛФА БЕТА")
    def test_invoice_normalization(self):self.assertEqual(normalize_invoice(" 000-12/AB "),"12AB")
    def test_sign_normalization(self):self.assertEqual(canonical_amount(-125.5),125.5)
    def test_exact_supplier_matching(self):
        m=match_suppliers([rec("SAP","S1","Алфа ООД",1)],[rec("AZHUR","A1","АЛФА ЕООД",1)],[],92,80)
        self.assertEqual(m[0]["match_method"],"EXACT_NORMALIZED_NAME")
    def test_fuzzy_threshold_is_conservative(self):
        m=match_suppliers([rec("SAP","S1","ALPHA INDUSTRIES",1)],[rec("AZHUR","A1","OMEGA SERVICES",1)],[],92,80)
        self.assertEqual(m[0]["match_method"],"UNMATCHED_SAP")
    def test_balance_tolerance(self):
        m=match_suppliers([rec("SAP","S","ALPHA",10.00)],[rec("AZHUR","A","ALPHA",10.04)],[],92,80)
        out=reconcile_suppliers([rec("SAP","S","ALPHA",10.00)],[rec("AZHUR","A","ALPHA",10.04)],m,.05)
        self.assertEqual(out[0]["financial_match"],"YES")
    def test_balance_outside_tolerance(self):
        m=match_suppliers([rec("SAP","S","ALPHA",10)],[rec("AZHUR","A","ALPHA",10.06)],[],92,80)
        self.assertEqual(reconcile_suppliers([rec("SAP","S","ALPHA",10)],[rec("AZHUR","A","ALPHA",10.06)],m,.05)[0]["exception_code"],"03_BALANCE_DIFFERENCE")
    def test_document_exact(self):
        s=[rec("SAP","S","ALPHA",10)];a=[rec("AZHUR","A","ALPHA",10)];m=match_suppliers(s,a,[],92,80);sr=reconcile_suppliers(s,a,m,.05)
        self.assertEqual(match_documents(s,a,sr,.05,5)[0]["document_match_method"],"EXACT_REFERENCE_AMOUNT")
    def test_document_date_amount_strong(self):
        s=[rec("SAP","S","ALPHA",10,"X")];a=[rec("AZHUR","A","ALPHA",10,"Y",dt.date(2026,1,4))];m=match_suppliers(s,a,[],92,80);sr=reconcile_suppliers(s,a,m,.05)
        self.assertEqual(match_documents(s,a,sr,.05,5)[0]["document_match_status"],"MATCHED")
    def test_document_unmatched_classification(self):
        s=[rec("SAP","S","ALPHA",10,"X")];a=[rec("AZHUR","A","ALPHA",12,"Y")];m=match_suppliers(s,a,[],92,80);sr=reconcile_suppliers(s,a,m,.05)
        self.assertEqual(match_documents(s,a,sr,.05,5)[0]["exception_code"],"04_DOCUMENT_DIFFERENCE")
if __name__=="__main__":unittest.main()
