from __future__ import annotations
import csv,datetime as dt,json,re,unicodedata
from collections import Counter,defaultdict
from difflib import SequenceMatcher
from pathlib import Path
from .xlsxio import read_workbook,write_workbook

FIELDS=["source_system","source_file","source_account","supplier_source_id","supplier_name_raw","supplier_name_normalized","supplier_key","gl_account","sap_scope_category","invoice_raw","invoice_normalized","document_number","reference","assignment","document_date","posting_date","due_date","currency","document_amount","company_currency_amount","debit_credit","signed_amount","text"]
EXCEPTIONS=["00_MATCH","01_SAP_ONLY","02_AZHUR_ONLY","03_BALANCE_DIFFERENCE","04_DOCUMENT_DIFFERENCE","05_CURRENCY_DIFFERENCE","06_TIMING_DIFFERENCE","07_SUPPLIER_MAPPING_DIFFERENCE","08_GL_CLASSIFICATION_DIFFERENCE","09_MIGRATION_OR_OPENING_BALANCE","10_ROUNDING_OR_FX","11_PARTIAL_PAYMENT_OR_RESIDUAL","12_MANUAL_REVIEW"]
LEGAL_FORMS={"ООД","ЕООД","АД","ЕАД","ЕТ"}

def normalize_supplier(value: object)->str:
    text=unicodedata.normalize("NFKC",str(value or "")).upper().replace("“",'"').replace("”",'"').replace("„",'"')
    text=re.sub(r"[^\w\s]"," ",text,flags=re.UNICODE);parts=text.split()
    while parts and parts[-1] in LEGAL_FORMS:parts.pop()
    return " ".join(parts)

def normalize_invoice(value: object)->str:
    text=unicodedata.normalize("NFKC",str(value or "")).upper().strip()
    text=re.sub(r"[\s./\\_-]+","",text)
    text=re.sub(r"^0+(?=\d)","",text)
    return text

def parse_date(value: object):
    if value in (None,""):return None
    text=str(value).strip()
    for fmt in ("%d.%m.%Y","%Y-%m-%d"):
        try:return dt.datetime.strptime(text,fmt).date()
        except ValueError:pass
    try:return dt.date(1899,12,30)+dt.timedelta(days=int(float(text)))
    except (ValueError,OverflowError):return None

def number(v):
    try:return float(str(v).replace(" ","").replace(",","."))
    except (ValueError,TypeError):return None

def canonical_amount(raw: float|None)->float|None:
    """Canonical convention: positive means a liability owed to the supplier."""
    return None if raw is None else -raw

def _blank():return {k:"" for k in FIELDS}

def load_azhur(paths):
    records=[];meta=[]
    for path in paths:
        books=read_workbook(path)
        for sheet,rows in books.items():
            if len(rows)<10:continue
            headers=rows[9];account="4012" if "401 / 2" in str(rows[3][0]) else "4011"
            for rownum,row in enumerate(rows[10:],11):
                row=row+[""]*(9-len(row));rec=_blank();raw=number(row[8] if account=="4012" else row[5])
                rec.update(source_system="AZHUR",source_file=Path(path).name,source_account=account,supplier_source_id=str(row[0]),supplier_name_raw=row[1],supplier_name_normalized=normalize_supplier(row[1]),gl_account=account,invoice_raw=row[2],invoice_normalized=normalize_invoice(row[2]),document_date=parse_date(row[3]),due_date=parse_date(row[6] if account=="4012" else row[4]),currency=(row[4] if account=="4012" else "EUR"),document_amount=number(row[7]) if account=="4012" else raw,company_currency_amount=raw,signed_amount=canonical_amount(raw))
                rec["_row"]=rownum;records.append(rec)
            meta.append({"file":Path(path).name,"sheet":sheet,"used_range":f"A1:{chr(64+len(headers))}{len(rows)}","header_row":10,"rows":len(rows)-10,"columns":headers,"account":account})
    return records,meta

def load_sap(path,scope):
    records=[];meta=[]
    for sheet,rows in read_workbook(path).items():
        headers=rows[0];detail=rows[2:] # row 2 is an embedded grand total
        for rownum,row in enumerate(detail,3):
            row=row+[""]*(18-len(row));rec=_blank();raw=number(row[9]);gl=str(row[4])
            rec.update(source_system="SAP",source_file=Path(path).name,supplier_source_id=str(row[6]),supplier_name_raw=row[7],supplier_name_normalized=normalize_supplier(row[7]),gl_account=gl,sap_scope_category=scope.get(gl,"OUT_OF_SCOPE"),invoice_raw=row[13] or row[16] or row[14],invoice_normalized=normalize_invoice(row[13] or row[16] or row[14]),document_number=row[14],reference=row[13],assignment=row[16],document_date=parse_date(row[12]),currency=row[11],document_amount=number(row[10]),company_currency_amount=raw,debit_credit=row[3],signed_amount=canonical_amount(raw),text=row[15]);rec["_row"]=rownum;records.append(rec)
        meta.append({"file":Path(path).name,"sheet":sheet,"used_range":f"A1:R{len(rows)}","header_row":1,"rows":len(detail),"embedded_total_row":2,"embedded_total":number(rows[1][9]),"columns":headers})
    return records,meta

def load_manual(path):
    p=Path(path)
    if not p.exists():return []
    with p.open(encoding="utf-8-sig",newline="") as f:return [r for r in csv.DictReader(f) if str(r.get("manual_override","")).lower() in {"1","true","yes","y"}]

def match_suppliers(sap,azhur,manual,fuzzy_auto=92,fuzzy_review=80):
    ss={r["supplier_source_id"]:r["supplier_name_raw"] for r in sap};aa={r["supplier_source_id"]:r["supplier_name_raw"] for r in azhur}
    mappings=[];used_s=set();used_a=set()
    for m in manual:
        sid=m["sap_vendor_id"];aid=m["azhur_supplier_id"]
        if sid in ss and aid in aa:mappings.append({"sap_vendor_id":sid,"sap_vendor_name":ss[sid],"azhur_supplier_id":aid,"azhur_supplier_name":aa[aid],"normalized_name":normalize_supplier(ss[sid]),"match_method":"MANUAL_OVERRIDE","match_score":100,"match_confidence":"CONFIRMED","manual_review_required":"NO"});used_s.add(sid);used_a.add(aid)
    byname=defaultdict(list)
    for aid,name in aa.items():byname[normalize_supplier(name)].append(aid)
    for sid,name in ss.items():
        if sid in used_s:continue
        candidates=[a for a in byname[normalize_supplier(name)] if a not in used_a]
        if len(candidates)==1:
            aid=candidates[0];mappings.append({"sap_vendor_id":sid,"sap_vendor_name":name,"azhur_supplier_id":aid,"azhur_supplier_name":aa[aid],"normalized_name":normalize_supplier(name),"match_method":"EXACT_NORMALIZED_NAME","match_score":100,"match_confidence":"HIGH","manual_review_required":"NO"});used_s.add(sid);used_a.add(aid)
    for sid,name in ss.items():
        if sid in used_s:continue
        scores=sorted(((SequenceMatcher(None,normalize_supplier(name),normalize_supplier(an)).ratio()*100,aid) for aid,an in aa.items() if aid not in used_a),reverse=True)
        score,aid=scores[0] if scores else (0,"")
        if score>=fuzzy_auto:
            mappings.append({"sap_vendor_id":sid,"sap_vendor_name":name,"azhur_supplier_id":aid,"azhur_supplier_name":aa[aid],"normalized_name":normalize_supplier(name),"match_method":"STRONG_FUZZY_NAME","match_score":round(score,1),"match_confidence":"HIGH","manual_review_required":"NO"});used_s.add(sid);used_a.add(aid)
        elif score>=fuzzy_review:
            mappings.append({"sap_vendor_id":sid,"sap_vendor_name":name,"azhur_supplier_id":"","azhur_supplier_name":aa[aid],"normalized_name":normalize_supplier(name),"match_method":"FUZZY_CANDIDATE","match_score":round(score,1),"match_confidence":"LOW","manual_review_required":"YES"});used_s.add(sid)
    for sid,name in ss.items():
        if sid not in used_s:mappings.append({"sap_vendor_id":sid,"sap_vendor_name":name,"azhur_supplier_id":"","azhur_supplier_name":"","normalized_name":normalize_supplier(name),"match_method":"UNMATCHED_SAP","match_score":0,"match_confidence":"NONE","manual_review_required":"YES"})
    for aid,name in aa.items():
        if aid not in used_a:mappings.append({"sap_vendor_id":"","sap_vendor_name":"","azhur_supplier_id":aid,"azhur_supplier_name":name,"normalized_name":normalize_supplier(name),"match_method":"UNMATCHED_AZHUR","match_score":0,"match_confidence":"NONE","manual_review_required":"YES"})
    return mappings

def reconcile_suppliers(sap,azhur,mappings,tol):
    sb=defaultdict(float);ab=defaultdict(float)
    for r in sap:sb[r["supplier_source_id"]]+=r["signed_amount"] or 0
    for r in azhur:ab[r["supplier_source_id"]]+=r["signed_amount"] or 0
    out=[]
    for m in mappings:
        sv,av=m["sap_vendor_id"],m["azhur_supplier_id"];s=sb[sv] if sv else 0;a=ab[av] if av else 0;diff=s-a
        financial="YES" if sv and av and abs(diff)<=tol else "NO"
        code="00_MATCH" if financial=="YES" else ("01_SAP_ONLY" if sv and not av else "02_AZHUR_ONLY" if av and not sv else "07_SUPPLIER_MAPPING_DIFFERENCE" if m["manual_review_required"]=="YES" else "03_BALANCE_DIFFERENCE")
        out.append({**m,"supplier_key":f'SAP:{sv}|AZHUR:{av}',"sap_balance":s,"azhur_balance":a,"difference":diff,"absolute_difference":abs(diff),"financial_match":financial,"document_match":"PENDING" if sv and av else "NO","accounting_match":"YES" if sv and av else "N/A","exception_code":code,"review_status":"OPEN" if code!="00_MATCH" else "RECONCILED","comment":""})
    return out

def match_documents(sap,azhur,supplier_recon,amount_tol,date_days):
    result=[]
    sapby=defaultdict(list);azhby=defaultdict(list)
    for r in sap:sapby[r["supplier_source_id"]].append(r)
    for r in azhur:azhby[r["supplier_source_id"]].append(r)
    for sr in supplier_recon:
        ss=sapby[sr["sap_vendor_id"]];aa=azhby[sr["azhur_supplier_id"]];used=set();matched=0
        for s in ss:
            candidates=[]
            for i,a in enumerate(aa):
                if i in used:continue
                delta=abs((s["signed_amount"] or 0)-(a["signed_amount"] or 0));same_ref=s["invoice_normalized"] and s["invoice_normalized"]==a["invoice_normalized"]
                dates=s["document_date"] and a["document_date"] and abs((s["document_date"]-a["document_date"]).days)<=date_days
                if delta<=amount_tol and (same_ref or (dates and s["currency"]==a["currency"])):candidates.append((0 if same_ref else 1,delta,i,a))
            if candidates:
                strength,delta,i,a=min(candidates,key=lambda x:(x[0],x[1]));used.add(i);matched+=abs(s["signed_amount"] or 0);method="EXACT_REFERENCE_AMOUNT" if strength==0 else "STRONG_AMOUNT_CURRENCY_DATE";score=100 if strength==0 else 90;code="00_MATCH"
            else:a={};delta=s["signed_amount"] or 0;method="UNMATCHED_SAP_DOCUMENT";score=0;code="04_DOCUMENT_DIFFERENCE"
            result.append(_docrow(sr,s,a,delta,method,score,code))
        for i,a in enumerate(aa):
            if i not in used:result.append(_docrow(sr,{},a,-(a["signed_amount"] or 0),"UNMATCHED_AZHUR_DOCUMENT",0,"04_DOCUMENT_DIFFERENCE"))
        sr["document_match"]="YES" if all(r["exception_code"]=="00_MATCH" for r in result if r["supplier_key"]==sr["supplier_key"]) else "NO"
        if sr["financial_match"]=="YES" and sr["document_match"]=="NO":sr["exception_code"]="04_DOCUMENT_DIFFERENCE";sr["review_status"]="OPEN"
    return result

def _docrow(sr,s,a,diff,method,score,code):
    return {"supplier_key":sr["supplier_key"],"sap_vendor":sr["sap_vendor_id"],"sap_document":s.get("document_number",""),"sap_reference":s.get("reference",""),"sap_assignment":s.get("assignment",""),"sap_date":s.get("document_date",""),"sap_currency":s.get("currency",""),"sap_amount":s.get("signed_amount",""),"azhur_id":sr["azhur_supplier_id"],"azhur_invoice":a.get("invoice_raw",""),"azhur_date":a.get("document_date",""),"azhur_due_date":a.get("due_date",""),"azhur_currency":a.get("currency",""),"azhur_amount":a.get("signed_amount",""),"difference":diff,"document_match_method":method,"document_match_score":score,"document_match_status":"MATCHED" if code=="00_MATCH" else "UNMATCHED","exception_code":code}
def data_quality(sap,azhur,mappings,settings,controls):
    checks=[]
    def add(name,status,count,detail):checks.append({"control":name,"status":status,"count":count,"detail":detail})
    for system,data in (("SAP",sap),("AZHUR",azhur)):
        add(f"{system}: null supplier names","PASS" if not (n:=sum(not r["supplier_name_raw"] for r in data)) else "FAIL",n,"Must be zero")
        add(f"{system}: missing amounts","PASS" if not (n:=sum(r["signed_amount"] is None for r in data)) else "FAIL",n,"Must be zero")
        add(f"{system}: invalid dates","PASS" if not (n:=sum(r["document_date"] is None for r in data)) else "WARN",n,"Missing/unparseable source dates")
        unsupported={r["currency"] for r in data if r["currency"] not in settings["supported_currencies"]}
        add(f"{system}: unsupported currencies","PASS" if not unsupported else "FAIL",len(unsupported),", ".join(unsupported))
        keys=[(r["supplier_source_id"],r["invoice_normalized"],r["signed_amount"]) for r in data if r["invoice_normalized"]]
        add(f"{system}: duplicate exact documents","WARN" if len(keys)!=len(set(keys)) else "PASS",len(keys)-len(set(keys)),"Potential split/residual items require review")
    unknown={r["gl_account"] for r in sap if r["sap_scope_category"]=="OUT_OF_SCOPE"}
    add("SAP: unknown G/L accounts","FAIL" if unknown else "PASS",len(unknown),", ".join(sorted(unknown)))
    impossible=sum((r["debit_credit"]=="H" and (number(r["company_currency_amount"]) or 0)>0) or (r["debit_credit"]=="S" and (number(r["company_currency_amount"]) or 0)<0) for r in sap)
    add("SAP: impossible sign combinations","FAIL" if impossible else "PASS",impossible,"H is expected negative and S positive in this extract")
    mpairs=[(m["sap_vendor_id"],m["azhur_supplier_id"]) for m in mappings if m["sap_vendor_id"] and m["azhur_supplier_id"]]
    add("Mapping: duplicate pairs","FAIL" if len(mpairs)!=len(set(mpairs)) else "PASS",len(mpairs)-len(set(mpairs)),"Must be zero")
    many=max(0,len(mpairs)-len({x[0] for x in mpairs}))+max(0,len(mpairs)-len({x[1] for x in mpairs}))
    add("Mapping: many-to-many suppliers","FAIL" if many else "PASS",many,"Each source ID maps once")
    add("Control: SAP normalization balance","PASS" if abs(controls["sap_raw_detail"]+controls["sap_canonical_all"])<0.005 else "FAIL",abs(controls["sap_raw_detail"]+controls["sap_canonical_all"]),"Canonical total must exactly reverse source sign")
    add("Control: Azhur normalization balance","PASS" if abs(controls["azhur_raw"]+controls["azhur_canonical"])<0.005 else "FAIL",abs(controls["azhur_raw"]+controls["azhur_canonical"]),"Canonical total must exactly reverse source sign")
    return checks

def _rows(records,columns):return [columns]+[[r.get(c,"") for c in columns] for r in records]

def run(config_path="config/settings.json"):
    settings=json.loads(Path(config_path).read_text(encoding="utf-8"));base=Path(config_path).resolve().parents[1];inp=settings["input"]
    azhur,ameta=load_azhur([base/p for p in inp["azhur_files"]]);sap_all,smeta=load_sap(base/inp["sap_file"],settings["sap_account_scope"])
    sap=[r for r in sap_all if r["sap_scope_category"] in settings["reconciliation_scope_categories"]]
    tol=settings["tolerances"];mappings=match_suppliers(sap,azhur,load_manual(base/settings["mapping_file"]),tol["fuzzy_auto_threshold"],tol["fuzzy_review_threshold"])
    supplier=reconcile_suppliers(sap,azhur,mappings,tol["supplier_balance_tolerance"]);docs=match_documents(sap,azhur,supplier,tol["document_amount_tolerance"],tol["date_tolerance_days"])
    controls={"sap_source_embedded_total":smeta[0]["embedded_total"],"sap_raw_detail":sum(r["company_currency_amount"] or 0 for r in sap_all),"sap_canonical_all":sum(r["signed_amount"] or 0 for r in sap_all),"sap_in_scope":sum(r["signed_amount"] or 0 for r in sap),"azhur_raw":sum(r["company_currency_amount"] or 0 for r in azhur),"azhur_canonical":sum(r["signed_amount"] or 0 for r in azhur)}
    quality=data_quality(sap_all,azhur,mappings,settings,controls)
    mapped=[r for r in supplier if r["sap_vendor_id"] and r["azhur_supplier_id"]];financial=[r for r in mapped if r["financial_match"]=="YES"]
    denom=sum(abs(r["sap_balance"]) for r in supplier if r["sap_vendor_id"]);reconvalue=sum(abs(r["sap_balance"]) for r in financial)
    metrics={"Run timestamp (UTC)":dt.datetime.now(dt.timezone.utc).isoformat(),"Source files":", ".join(inp["azhur_files"]+[inp["sap_file"]]),"SAP total AP balance (in scope)":controls["sap_in_scope"],"Azhur total AP balance":controls["azhur_canonical"],"Net difference (SAP - Azhur)":controls["sap_in_scope"]-controls["azhur_canonical"],"Reconciliation coverage %":reconvalue/denom if denom else 0,"Supplier population SAP":len({r["supplier_source_id"] for r in sap}),"Supplier population Azhur":len({r["supplier_source_id"] for r in azhur}),"Mapped suppliers":len(mapped),"Unmatched SAP suppliers":sum(r["match_method"]=="UNMATCHED_SAP" for r in mappings),"Unmatched Azhur suppliers":sum(r["match_method"]=="UNMATCHED_AZHUR" for r in mappings),"Financially reconciled suppliers %":len(financial)/len(mapped) if mapped else 0,"Financially reconciled balance %":reconvalue/denom if denom else 0,"Document-reconciled balance %":sum(abs(r["sap_balance"]) for r in mapped if r["document_match"]=="YES")/denom if denom else 0,"Exceptions count":sum(r["exception_code"]!="00_MATCH" for r in supplier),"Material exceptions count":sum(r["absolute_difference"]>=tol["material_exception_threshold"] for r in supplier),"Absolute unresolved difference":sum(r["absolute_difference"] for r in supplier if r["exception_code"]!="00_MATCH")}
    exceptions=[]
    for r in sorted((x for x in supplier if x["exception_code"]!="00_MATCH"),key=lambda x:x["absolute_difference"],reverse=True):
        action="Confirm supplier mapping" if r["exception_code"]=="07_SUPPLIER_MAPPING_DIFFERENCE" else "Investigate unmatched source documents and cut-off"
        exceptions.append({**r,"root_cause":"Unresolved source population or composition difference","suggested_action":action,"finance_comment":""})
    controlrows=[["Metric","Value","Definition"]]+[[k,v,"Value-weighted coverage = absolute SAP balance of financially matched suppliers / absolute total in-scope SAP balance" if "coverage" in k.lower() else ""] for k,v in metrics.items()]
    control_detail=[{"control":"Source metadata","status":"INFO","count":sum(x["rows"] for x in smeta+ameta),"detail":json.dumps(smeta+ameta,ensure_ascii=False)}]+[{"control":k,"status":"INFO","count":v,"detail":"Transformation control total"} for k,v in controls.items()]+quality
    supcols=["supplier_key","sap_vendor_id","sap_vendor_name","azhur_supplier_id","azhur_supplier_name","match_method","match_score","match_confidence","manual_review_required","sap_balance","azhur_balance","difference","absolute_difference","financial_match","document_match","accounting_match","exception_code","review_status","comment"]
    doccols=list(docs[0]) if docs else ["supplier_key"]
    excols=supcols+["root_cause","suggested_action","finance_comment"]
    mapcols=["sap_vendor_id","sap_vendor_name","azhur_supplier_id","azhur_supplier_name","normalized_name","match_method","match_score","match_confidence","manual_review_required"]
    normcols=FIELDS
    sheets=[("0. Control",controlrows),("1. Supplier Recon",_rows(supplier,supcols)),("2. Document Recon",_rows(docs,doccols)),("3. Exceptions",_rows(exceptions,excols)),("4. Supplier Mapping",_rows(mappings,mapcols)),("5. SAP Normalized",_rows(sap_all,normcols)),("6. Azhur Normalized",_rows(azhur,normcols)),("7. Controls",_rows(control_detail,["control","status","count","detail"]))]
    output=base/settings["output_file"];write_workbook(output,sheets,{"Value","sap_balance","azhur_balance","difference","absolute_difference","sap_amount","azhur_amount","signed_amount","document_amount","company_currency_amount"})
    runlog={"run_timestamp":metrics["Run timestamp (UTC)"],"input_filenames":inp["azhur_files"]+[inp["sap_file"]],"source_metadata":ameta+smeta,"source_totals":controls,"mapped_suppliers":len(mapped),"auto_matches":sum(m["match_method"] in {"EXACT_NORMALIZED_NAME","STRONG_FUZZY_NAME"} for m in mappings),"manual_mappings":sum(m["match_method"]=="MANUAL_OVERRIDE" for m in mappings),"exception_counts":Counter(r["exception_code"] for r in supplier),"output_filename":settings["output_file"],"metrics":metrics,"control_failures":[q for q in quality if q["status"]=="FAIL"]}
    Path(base/settings["run_log"]).write_text(json.dumps(runlog,ensure_ascii=False,indent=2,default=str),encoding="utf-8")
    return runlog
