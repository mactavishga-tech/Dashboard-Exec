from zipfile import ZipFile, ZIP_DEFLATED
from html import escape
from datetime import date, timedelta

OUT='The_Harlow_Hotel_Executive_Dashboard_Visual_Prototype.xlsm'
NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships'

def col(n):
    s=''
    while n:
        n,r=divmod(n-1,26); s=chr(65+r)+s
    return s

def cell_ref(r,c): return f'{col(c)}{r}'

def c_xml(r,c,v,style=0):
    ref=cell_ref(r,c)
    if isinstance(v,(int,float)):
        return f'<c r="{ref}" s="{style}"><v>{v}</v></c>'
    if isinstance(v,str) and v.startswith('='):
        return f'<c r="{ref}" s="{style}"><f>{escape(v[1:])}</f></c>'
    return f'<c r="{ref}" t="inlineStr" s="{style}"><is><t>{escape(str(v))}</t></is></c>'

def sheet_xml(rows, merges=(), col_widths=(), tables=()):
    maxr=max(rows) if rows else 1; maxc=max((max(v) for v in rows.values() if v), default=1)
    cols=''.join(f'<col min="{a}" max="{b}" width="{w}" customWidth="1"/>' for a,b,w in col_widths)
    data=[]
    for r in range(1,maxr+1):
        cells=''.join(c_xml(r,c,*rows[r][c]) for c in sorted(rows.get(r,{})))
        if cells: data.append(f'<row r="{r}" ht="26" customHeight="1">{cells}</row>')
    merge_xml=f'<mergeCells count="{len(merges)}">' + ''.join(f'<mergeCell ref="{m}"/>' for m in merges) + '</mergeCells>' if merges else ''
    tableparts=f'<tableParts count="{len(tables)}">' + ''.join(f'<tablePart r:id="rId{i+1}"/>' for i,_ in enumerate(tables)) + '</tableParts>' if tables else ''
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{NS}" xmlns:r="{REL}"><sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="18"/><cols>{cols}</cols><sheetData>{"".join(data)}</sheetData>{merge_xml}<pageMargins left="0.25" right="0.25" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>{tableparts}</worksheet>'

def table_xml(tid,name,ref,headers):
    cols=''.join(f'<tableColumn id="{i}" name="{escape(h)}"/>' for i,h in enumerate(headers,1))
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="{NS}" id="{tid}" name="{name}" displayName="{name}" ref="{ref}" totalsRowShown="0"><autoFilter ref="{ref}"/><tableColumns count="{len(headers)}">{cols}</tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>'

# styles: 0 normal, 1 title, 2 subtitle, 3 dark header, 4 gold kpi, 5 blue kpi, 6 red alert, 7 green, 8 small, 9 money, 10 percent, 11 table header, 12 bar
styles='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="2"><numFmt numFmtId="164" formatCode="£#,##0"/><numFmt numFmtId="165" formatCode="0.0%"/></numFmts><fonts count="6"><font><sz val="10"/><color rgb="FF1F2937"/><name val="Aptos"/></font><font><b/><sz val="24"/><color rgb="FFFFFFFF"/><name val="Aptos Display"/></font><font><sz val="11"/><color rgb="FFE5E7EB"/><name val="Aptos"/></font><font><b/><sz val="12"/><color rgb="FFFFFFFF"/><name val="Aptos"/></font><font><b/><sz val="18"/><color rgb="FF111827"/><name val="Aptos"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Aptos"/></font></fonts><fills count="9"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF0B1F33"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFC8A45D"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF8FAFC"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF0F4C81"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFB91C1C"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF166534"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF374151"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFD1D5DB"/></left><right style="thin"><color rgb="FFD1D5DB"/></right><top style="thin"><color rgb="FFD1D5DB"/></top><bottom style="thin"><color rgb="FFD1D5DB"/></bottom></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="13"><xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFill="1" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFill="1" applyFont="1" applyAlignment="1"><alignment horizontal="center"/></xf><xf numFmtId="0" fontId="3" fillId="5" borderId="1" xfId="0" applyFill="1" applyFont="1" applyAlignment="1"><alignment horizontal="center"/></xf><xf numFmtId="0" fontId="4" fillId="3" borderId="1" xfId="0" applyFill="1" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="4" fillId="4" borderId="1" xfId="0" applyFill="1" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="3" fillId="6" borderId="1" xfId="0" applyFill="1" applyFont="1" applyAlignment="1"><alignment horizontal="center" wrapText="1"/></xf><xf numFmtId="0" fontId="3" fillId="7" borderId="1" xfId="0" applyFill="1" applyFont="1"/><xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1"/><xf numFmtId="164" fontId="4" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyAlignment="1"><alignment horizontal="center"/></xf><xf numFmtId="165" fontId="4" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyAlignment="1"><alignment horizontal="center"/></xf><xf numFmtId="0" fontId="5" fillId="8" borderId="1" xfId="0" applyFill="1" applyFont="1"/><xf numFmtId="0" fontId="3" fillId="2" borderId="1" xfId="0" applyFill="1" applyFont="1" applyAlignment="1"><alignment horizontal="center" wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="1" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>'''

rows={}
def put(r,c,v,s=0): rows.setdefault(r,{})[c]=(v,s)
for c in range(1,17): put(1,c,'',1); put(2,c,'',2)
put(1,1,'THE HARLOW HOTEL',1); put(2,1,'Executive Dashboard Visual Prototype | GBP | Twin & Double Room Types Only',2)
for label,val,st,c in [('Occupancy','=AVERAGE(Placeholder_Data!C5:C18)',10,1),('ADR','=AVERAGE(Placeholder_Data!D5:D18)',9,4),('RevPAR','=AVERAGE(Placeholder_Data!E5:E18)',9,7),('RPS','=AVERAGE(Placeholder_Data!F5:F18)',9,10),('Total Revenue','=SUM(Placeholder_Data!G5:G18)',9,13)]:
    put(4,c,label,3); put(5,c,val,st)
sections=[(7,1,'Daily Operations Strip'),(7,7,'Executive Alerts'),(7,12,'Data Health'),(14,1,'Room Performance Score (RPS)')]
for r,c,t in sections: put(r,c,t,11)
for i,h in enumerate(['Date','Arrivals','Departures','Stayovers','OOO','Notes'],1): put(8,i,h,3)
for i,row in enumerate([['27 Jun 2026',34,29,88,2,'Weekend leisure peak'],['28 Jun 2026',22,37,74,1,'Late checkout pressure'],['29 Jun 2026',41,18,96,3,'Corporate block arrivals']],9):
    for c,v in enumerate(row,1): put(i,c,v,8)
for i,h in enumerate(['Severity','Alert','Owner','Action'],7): put(8,i,h,3)
alerts=[['High','Double ADR below target','Revenue Manager','Review fenced offers'],['Medium','Twin housekeeping turn risk','Ops Manager','Add 2 PM attendants'],['Low','RPS survey sample thin','GM','Monitor next 48h']]
for r,row in enumerate(alerts,9):
    for c,v in enumerate(row,7): put(r,c,v,6 if c==7 and v=='High' else 8)
for i,h in enumerate(['Metric','Status','Last Refresh','Notes'],12): put(8,i,h,3)
for r,row in enumerate([['Source tables','Healthy','27 Jun 2026 07:00','All tables populated'],['Room types','Healthy','27 Jun 2026 07:00','Twin/Double only'],['Currency','Healthy','27 Jun 2026 07:00','GBP']],9):
    for c,v in enumerate(row,12): put(r,c,v,7 if c==13 else 8)
for i,h in enumerate(['Room Type','Occupancy','ADR','RevPAR','Guest Score','RPS'],1): put(15,i,h,3)
for r,row in enumerate([['Twin','=AVERAGEIF(Placeholder_Data!B5:B18,"Twin",Placeholder_Data!C5:C18)','=AVERAGEIF(Placeholder_Data!B5:B18,"Twin",Placeholder_Data!D5:D18)','=AVERAGEIF(Placeholder_Data!B5:B18,"Twin",Placeholder_Data!E5:E18)','=AVERAGEIF(Placeholder_Data!B5:B18,"Twin",Placeholder_Data!H5:H18)','=AVERAGEIF(Placeholder_Data!B5:B18,"Twin",Placeholder_Data!F5:F18)'],['Double','=AVERAGEIF(Placeholder_Data!B5:B18,"Double",Placeholder_Data!C5:C18)','=AVERAGEIF(Placeholder_Data!B5:B18,"Double",Placeholder_Data!D5:D18)','=AVERAGEIF(Placeholder_Data!B5:B18,"Double",Placeholder_Data!E5:E18)','=AVERAGEIF(Placeholder_Data!B5:B18,"Double",Placeholder_Data!H5:H18)','=AVERAGEIF(Placeholder_Data!B5:B18,"Double",Placeholder_Data!F5:F18)']],16):
    for c,v in enumerate(row,1): put(r,c,v,10 if c==2 else (9 if c in [3,4,6] else 8))
put(20,1,'Key Insights',12); put(21,1,'• Double rooms are the strongest revenue driver this week.  • Twin turn pressure is operationally visible across arrivals/departures.  • Data health checks confirm GBP currency and Twin/Double-only room mix.',12)
merges=['A1:P1','A2:P2','A4:C4','A5:C6','D4:F4','D5:F6','G4:I4','G5:I6','J4:L4','J5:L6','M4:P4','M5:P6','A7:F7','G7:K7','L7:P7','A14:F14','A20:P20','A21:P22']

# data sheet
headers=['Date','RoomType','Occupancy','ADR_GBP','RevPAR_GBP','RPS_GBP','Revenue_GBP','GuestScore','Arrivals','Departures','Stayovers','OOO_Rooms']
drows={1:{1:('The Harlow Hotel Placeholder Data Tables',1)},3:{}}
for c,h in enumerate(headers,1): drows[4]={**drows.get(4,{}), c:(h,11)}
base=date(2026,6,20)
for i in range(14):
    rt='Twin' if i%2==0 else 'Double'; occ=.72+(i%5)*.035; adr=145+(i%4)*12+(10 if rt=='Double' else 0); rev=occ*adr; rps=rev*(.92+(i%3)*.03); revenue=round(rev*(42 if rt=='Twin' else 58))
    vals=[(base+timedelta(days=i)).isoformat(),rt,round(occ,3),adr,round(rev,2),round(rps,2),revenue,round(4.2+(i%4)*.12,2),18+i%9,16+i%8,70+i%25,i%4]
    for c,v in enumerate(vals,1): drows.setdefault(5+i,{})[c]=(v,10 if c==3 else (9 if c in [4,5,6,7] else 8))
# extra tables lower
for c,h in enumerate(['AlertID','Severity','Alert','Owner','Action','Status'],1): drows.setdefault(22,{})[c]=(h,11)
for r,vals in enumerate([[1,'High','Double ADR below target','Revenue Manager','Review fenced offers','Open'],[2,'Medium','Twin housekeeping turn risk','Ops Manager','Add 2 PM attendants','Open'],[3,'Low','RPS survey sample thin','GM','Monitor next 48h','Watching']],23):
    for c,v in enumerate(vals,1): drows.setdefault(r,{})[c]=(v,8)
for c,h in enumerate(['CheckID','CheckName','Status','LastRefresh','Notes'],8): drows.setdefault(22,{})[c]=(h,11)
for r,vals in enumerate([[1,'Source tables','Healthy','2026-06-27 07:00','All tables populated'],[2,'Room types','Healthy','2026-06-27 07:00','Twin/Double only'],[3,'Currency','Healthy','2026-06-27 07:00','GBP']],23):
    for c,v in enumerate(vals,8): drows.setdefault(r,{})[c]=(v,8)
for c,h in enumerate(['InsightID','Insight'],1): drows.setdefault(29,{})[c]=(h,11)
for r,vals in enumerate([[1,'Double rooms are the strongest revenue driver this week.'],[2,'Twin turn pressure is operationally visible across arrivals/departures.'],[3,'Data health checks confirm GBP currency and Twin/Double-only room mix.']],30):
    for c,v in enumerate(vals,1): drows.setdefault(r,{})[c]=(v,8)

for c,h in enumerate(['Date','Arrivals','Departures','Stayovers','OOO_Rooms','Notes'],1): drows.setdefault(35,{})[c]=(h,11)
for r,vals in enumerate([['2026-06-27',34,29,88,2,'Weekend leisure peak'],['2026-06-28',22,37,74,1,'Late checkout pressure'],['2026-06-29',41,18,96,3,'Corporate block arrivals']],36):
    for c,v in enumerate(vals,1): drows.setdefault(r,{})[c]=(v,8)

docprops='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application></Properties>'''
core='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>The Harlow Hotel Executive Dashboard Visual Prototype</dc:title><dc:creator>OpenAI Codex</dc:creator></cp:coreProperties>'''
workbook=f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="{NS}" xmlns:r="{REL}"><sheets><sheet name="Executive Dashboard" sheetId="1" r:id="rId1"/><sheet name="Placeholder_Data" sheetId="2" r:id="rId2"/></sheets><definedNames><definedName name="DashboardCurrency">"GBP"</definedName><definedName name="AllowedRoomTypes">Placeholder_Data!$B$5:$B$18</definedName></definedNames></workbook>'''
rels='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'''
wb_rels=f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{REL}/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="{REL}/styles" Target="styles.xml"/></Relationships>'''
content='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/><Override PartName="/xl/tables/table2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/><Override PartName="/xl/tables/table3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/><Override PartName="/xl/tables/table4.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/><Override PartName="/xl/tables/table5.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>'''
sheet2_rels=f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL}/table" Target="../tables/table1.xml"/><Relationship Id="rId2" Type="{REL}/table" Target="../tables/table2.xml"/><Relationship Id="rId3" Type="{REL}/table" Target="../tables/table3.xml"/><Relationship Id="rId4" Type="{REL}/table" Target="../tables/table4.xml"/><Relationship Id="rId5" Type="{REL}/table" Target="../tables/table5.xml"/></Relationships>'''
with ZipFile(OUT,'w',ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml',content); z.writestr('_rels/.rels',rels); z.writestr('docProps/app.xml',docprops); z.writestr('docProps/core.xml',core)
    z.writestr('xl/workbook.xml',workbook); z.writestr('xl/_rels/workbook.xml.rels',wb_rels); z.writestr('xl/styles.xml',styles)
    z.writestr('xl/worksheets/sheet1.xml',sheet_xml(rows,merges,[(1,16,13)],[]))
    z.writestr('xl/worksheets/sheet2.xml',sheet_xml(drows,['A1:L1'],[(1,12,15)],[(1,'tblRoomPerformance'),(2,'tblExecutiveAlerts'),(3,'tblDataHealth'),(4,'tblKeyInsights'),(5,'tblDailyOperations')]))
    z.writestr('xl/worksheets/_rels/sheet2.xml.rels',sheet2_rels)
    z.writestr('xl/tables/table1.xml',table_xml(1,'tblRoomPerformance','A4:L18',headers))
    z.writestr('xl/tables/table2.xml',table_xml(2,'tblExecutiveAlerts','A22:F25',['AlertID','Severity','Alert','Owner','Action','Status']))
    z.writestr('xl/tables/table3.xml',table_xml(3,'tblDataHealth','H22:L25',['CheckID','CheckName','Status','LastRefresh','Notes']))
    z.writestr('xl/tables/table4.xml',table_xml(4,'tblKeyInsights','A29:B32',['InsightID','Insight']))
    z.writestr('xl/tables/table5.xml',table_xml(5,'tblDailyOperations','A35:F38',['Date','Arrivals','Departures','Stayovers','OOO_Rooms','Notes']))
print(OUT)
