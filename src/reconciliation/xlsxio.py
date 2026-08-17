"""Small dependency-free OOXML reader/writer used in restricted finance environments."""
from __future__ import annotations
import datetime as dt
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

MAIN="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS={"m":MAIN,"r":REL}

def _col_index(ref: str) -> int:
    letters=re.match(r"[A-Z]+",ref).group()
    n=0
    for char in letters:n=n*26+ord(char)-64
    return n-1

def col_letter(n: int) -> str:
    result=""
    while n:
        n,rem=divmod(n-1,26);result=chr(65+rem)+result
    return result

def read_workbook(path: str|Path) -> dict[str,list[list[object]]]:
    """Read cell values and sheet names without evaluating formulas."""
    with zipfile.ZipFile(path) as archive:
        shared=[]
        if "xl/sharedStrings.xml" in archive.namelist():
            root=ET.fromstring(archive.read("xl/sharedStrings.xml"))
            shared=["".join(n.text or "" for n in item.iter() if n.tag.endswith("}t")) for item in root]
        workbook=ET.fromstring(archive.read("xl/workbook.xml"))
        relroot=ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rels={n.attrib["Id"]:n.attrib["Target"] for n in relroot}
        result={}
        for sheet in workbook.find("m:sheets",NS):
            target=rels[sheet.attrib[f"{{{REL}}}id"]]
            target=target[1:] if target.startswith("/xl/") else "xl/"+target.lstrip("/")
            root=ET.fromstring(archive.read(target));rows=[]
            for row in root.findall(".//m:sheetData/m:row",NS):
                cells={}
                for cell in row.findall("m:c",NS):
                    kind=cell.attrib.get("t");node=cell.find("m:v",NS)
                    value="" if node is None else node.text or ""
                    if kind=="s" and value:value=shared[int(value)]
                    elif kind=="inlineStr":value="".join(n.text or "" for n in cell.iter() if n.tag.endswith("}t"))
                    cells[_col_index(cell.attrib["r"])]=value
                rows.append([cells.get(i,"") for i in range(max(cells,default=-1)+1)])
            result[sheet.attrib["name"]]=rows
        return result

def _cell(value: object, ref: str, style: int=0) -> str:
    attr=f' r="{ref}" s="{style}"'
    if value is None or value=="":return f"<c{attr} t=\"inlineStr\"><is><t></t></is></c>"
    if isinstance(value,bool):return f"<c{attr} t=\"b\"><v>{int(value)}</v></c>"
    if isinstance(value,(int,float)):return f"<c{attr}><v>{value}</v></c>"
    if isinstance(value,(dt.date,dt.datetime)):
        serial=(value.date() if isinstance(value,dt.datetime) else value)-dt.date(1899,12,30)
        return f'<c r="{ref}" s="3"><v>{serial.days}</v></c>'
    return f'<c{attr} t="inlineStr"><is><t xml:space="preserve">{escape(str(value))}</t></is></c>'

def write_workbook(path: str|Path, sheets: list[tuple[str,list[list[object]]]], amount_headers: set[str]|None=None) -> None:
    """Write a styled, frozen, filtered audit workbook."""
    amount_headers=amount_headers or set();path=Path(path);path.parent.mkdir(parents=True,exist_ok=True)
    content=['<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>']
    for i in range(len(sheets)):content.append(f'<Override PartName="/xl/worksheets/sheet{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    content.append('</Types>')
    rootrels='<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    wb='<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="%s" xmlns:r="%s"><sheets>'%(MAIN,REL)+''.join(f'<sheet name="{escape(n)}" sheetId="{i+1}" r:id="rId{i+1}"/>' for i,(n,_) in enumerate(sheets))+'</sheets></workbook>'
    rels='<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+''.join(f'<Relationship Id="rId{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i+1}.xml"/>' for i in range(len(sheets)))+f'<Relationship Id="rId{len(sheets)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
    styles='''<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="2"><numFmt numFmtId="164" formatCode="#,##0.00;[Red]-#,##0.00"/><numFmt numFmtId="165" formatCode="yyyy-mm-dd"/></numFmts><fonts count="2"><font><sz val="10"/><name val="Aptos"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="10"/><name val="Aptos"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFill="1" applyFont="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'''
    with zipfile.ZipFile(path,"w",zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml","".join(content));z.writestr("_rels/.rels",rootrels);z.writestr("xl/workbook.xml",wb);z.writestr("xl/_rels/workbook.xml.rels",rels);z.writestr("xl/styles.xml",styles)
        for idx,(_,rows) in enumerate(sheets,1):
            width=max((len(r) for r in rows),default=1);height=max(len(rows),1);headers=rows[0] if rows else []
            cols=''.join(f'<col min="{i}" max="{i}" width="{min(42,max(12,max((len(str(r[i-1])) for r in rows[:500] if len(r)>=i),default=8)+2))}" customWidth="1"/>' for i in range(1,width+1))
            body=[]
            for rn,row in enumerate(rows,1):
                cells=[]
                for cn,val in enumerate(row,1):
                    style=1 if rn==1 else (2 if cn<=len(headers) and str(headers[cn-1]) in amount_headers else 0)
                    cells.append(_cell(val,f'{col_letter(cn)}{rn}',style))
                body.append(f'<row r="{rn}">{"".join(cells)}</row>')
            ref=f'A1:{col_letter(width)}{height}'
            cf=f'<conditionalFormatting sqref="A2:{col_letter(width)}{height}"><cfRule type="expression" priority="1"><formula>ISNUMBER(SEARCH("MANUAL",$A2&amp;$B2&amp;$C2&amp;$D2))</formula></cfRule></conditionalFormatting>'
            xml=f'<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{MAIN}"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>{cols}</cols><sheetData>{"".join(body)}</sheetData><autoFilter ref="{ref}"/>{cf}<pageMargins left="0.3" right="0.3" top="0.5" bottom="0.5" header="0.2" footer="0.2"/></worksheet>'
            z.writestr(f"xl/worksheets/sheet{idx}.xml",xml)
