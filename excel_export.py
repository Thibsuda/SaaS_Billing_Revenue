import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.legend import Legend
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.formatting.rule import FormulaRule

import io
import pandas as pd


def _dt(v):
    return None if pd.isna(v) else pd.Timestamp(v).to_pydatetime()


def to_lists(hub, stp):
    deals = [dict(id=r["Deal ID"], cust=r["Customer"], stage=r["Stage"], start=_dt(r["Start Date"]), val=float(r["Contract Value"]), end=_dt(r["End Date"]), prod=r["Product"]) for _, r in hub.iterrows()]
    inv = [dict(id=r["Invoice ID"], cust=r["Customer"], date=_dt(r["Invoice Date"]), due=_dt(r["Due Date"]), amt=float(r["Amount"]), paid=float(r["Amount Paid"]), status=r["Status"], deal=(r["Deal ID"] or None)) for _, r in stp.iterrows()]
    return deals, inv


def build_workbook(hub, stp, asof, first_month):
    """Return the workbook (with live formulas) as bytes."""
    deals, inv = to_lists(hub, stp)
    asof = pd.Timestamp(asof).to_pydatetime(); first_month = pd.Timestamp(first_month).to_pydatetime()
    ASOF = asof
    HLAST, SLAST, RLAST = max(200, len(deals) + 60), max(400, len(inv) + 100), max(202, len(deals) + 22)

    NAVY="1F2A44"; TEAL="0F766E"; GREY="5B6573"; SOFT="DCE6F1"; HELP="EEF3F9"; YEL="FFF2A8"; WHT="FFFFFF"; LG="F3F5F8"
    F="Arial"
    fn=lambda **k: Font(name=F,size=k.pop("size",10),**k)
    fill=lambda c: PatternFill("solid",fgColor=c)
    thin=Side(style="thin",color="D0D6DC"); BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
    CUR='$#,##0;($#,##0);-'; CUR2='$#,##0.00;($#,##0.00);-'; NUM='#,##0;(#,##0);-'; PCT='0.0%'; DT='yyyy-mm-dd'; MON='mmm yy'

    wb=Workbook(); SD=wb.active; SD.title="Summary_Dashboard"
    HW=wb.create_sheet("HubSpot_Working"); SW=wb.create_sheet("Stripe_Working"); RR=wb.create_sheet("RevRec_Waterfall"); DR=wb.create_sheet("Deferred_Rev")
    RM=wb.create_sheet("Read_Me"); RH=wb.create_sheet("Raw_HubSpot"); RS=wb.create_sheet("Raw_Stripe")

    def head(ws,row,col,labels,dark=NAVY,fc=WHT,h=30):
        for i,t in enumerate(labels):
            c=ws.cell(row=row,column=col+i,value=t); c.font=fn(bold=True,color=fc); c.fill=fill(dark)
            c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); c.border=BOX
        ws.row_dimensions[row].height=h
    def put(ws,ref,v,fmt=None,bold=False,inp=False,fl=None,color=None,italic=False,align=None,wrap=False):
        c=ws[ref]; c.value=v; c.font=fn(bold=bold,color=("0000FF" if inp else (color or "000000")),italic=italic); c.border=BOX
        if fmt: c.number_format=fmt
        if fl: c.fill=fill(fl)
        if align or wrap: c.alignment=Alignment(horizontal=align,wrap_text=wrap,vertical="center")
        return c
    def widths(ws,w):
        for k,v in w.items(): ws.column_dimensions[k].width=v
    def sect(ws,ref,t): ws[ref]=t; ws[ref].font=fn(bold=True,size=12,color=NAVY)
    def okfmt(ws,rng,first):
        ws.conditional_formatting.add(rng,FormulaRule(formula=[f'{first}="Ties"'],fill=PatternFill("solid",bgColor="D9EEDF",fgColor="D9EEDF")))
        ws.conditional_formatting.add(rng,FormulaRule(formula=[f'{first}="Check"'],fill=PatternFill("solid",bgColor="F8D7D3",fgColor="F8D7D3")))

    # ---------------- raw data ----------------
    hcols=["Deal ID","Customer","Stage","Start Date","Contract Value","End Date","Product"]
    scols=["Invoice ID","Customer","Invoice Date","Due Date","Amount","Amount Paid","Status","Deal ID"]
    hrows=[(d["id"],d["cust"],d["stage"],d["start"],d["val"],d["end"],d["prod"]) for d in deals]
    srows=[(x["id"],x["cust"],x["date"],x["due"],x["amt"],x["paid"],x["status"],x["deal"] or None) for x in inv]
    hf=[None,None,None,DT,CUR2,DT,None]; sf=[None,None,DT,DT,CUR2,CUR2,None,None]
    def load(ws,cols,rows,fmts,last,inp):
        head(ws,1,1,cols,dark=NAVY if inp else GREY)
        for r in range(2,last+1):
            for j,f in enumerate(fmts):
                c=ws.cell(row=r,column=j+1); c.font=fn(color="0000FF" if inp else "000000"); c.border=BOX
                if f: c.number_format=f
        for i,row in enumerate(rows):
            for j,v in enumerate(row): ws.cell(row=2+i,column=j+1,value=v)
    load(RH,hcols,hrows,hf,len(hrows)+1,False); load(RS,scols,srows,sf,len(srows)+1,False)
    load(HW,hcols,hrows,hf,HLAST,True); load(SW,scols,srows,sf,SLAST,True); SW["G1"].value="Stripe Status"
    widths(RH,{"A":10,"B":26,"C":13,"D":12,"E":15,"F":12,"G":28}); widths(RS,{"A":12,"B":26,"C":13,"D":12,"E":13,"F":13,"G":10,"H":10})

    # ---------------- Stripe_Working helpers ----------------
    head(SW,1,9,["Clean Customer","Duplicate Check","Valid Flag","Open Balance","Days Past Due","Aging Bucket","Status","HubSpot Contract Amount","HubSpot Stage","Invoice Check"],dark=SOFT,fc=NAVY)
    G='=IF($A{r}="","",'
    sform={
     "I":G+'TRIM(LOWER(B{r})))',
     "J":G+'IF(AND(G{r}<>"Void",COUNTIFS(I$2:I{r},I{r},E$2:E{r},E{r},H$2:H{r},H{r},G$2:G{r},"<>Void",C$2:C{r},">="&DATE(YEAR(C{r}),MONTH(C{r}),1),C$2:C{r},"<="&EOMONTH(C{r},0))>1),"Duplicate",""))',
     "K":G+'IF(OR(G{r}="Void",G{r}="Draft",J{r}="Duplicate"),"Exclude","Valid"))',
     "L":G+'IF(K{r}="Valid",E{r}-F{r},0))',
     "M":G+'IF(L{r}>0,MAX(0,Summary_Dashboard!$B$2-D{r}),0))',
     "N":G+'IF(K{r}<>"Valid","Excluded",IF(L{r}<=0,"Paid",IF(M{r}=0,"Current",IF(M{r}<=30,"1 to 30",IF(M{r}<=60,"31 to 60",IF(M{r}<=90,"61 to 90","Over 90")))))))',
     "O":G+'IF(K{r}<>"Valid",IF(G{r}="Void","Void",IF(G{r}="Draft","Draft","Duplicate")),IF(L{r}<=0,"Paid",IF(F{r}>0,"Partially paid",IF(M{r}>0,"Past due","Open")))))',
     "P":G+'IFERROR(INDEX(HubSpot_Working!$E:$E,MATCH(H{r},HubSpot_Working!$A:$A,0)),"No Matching Deal"))',
     "Q":G+'IFERROR(INDEX(HubSpot_Working!$C:$C,MATCH(H{r},HubSpot_Working!$A:$A,0)),"No Matching Deal"))',
     "R":G+'IF(K{r}<>"Valid","Excluded",IF(H{r}="","Missing Deal ID",IF(Q{r}="No Matching Deal","No matching deal",IF(Q{r}<>"Closed Won","Deal not won","OK")))))'}
    sfmt={"L":CUR2,"M":"0","P":CUR2}
    for r in range(2,SLAST+1):
        for c,f in sform.items():
            x=SW[f"{c}{r}"]; x.value=f.replace("{r}",str(r)); x.font=fn(); x.fill=fill(HELP); x.border=BOX
            if c in sfmt: x.number_format=sfmt[c]
    widths(SW,{"A":12,"B":26,"C":13,"D":12,"E":13,"F":13,"G":10,"H":10,"I":26,"J":15,"K":11,"L":13,"M":11,"N":14,"O":15,"P":16,"Q":15,"R":17}); SW.freeze_panes="C2"

    # ---------------- HubSpot_Working helpers ----------------
    head(HW,1,8,["Stripe Billed","Unbilled Variance","Billing Status","Contract Days","Revenue Recognized To Date","Billed Less Recognized","Won #"],dark=SOFT,fc=NAVY)
    G='=IF($A{r}="","",'
    hform={
     "H":G+'SUMIFS(Stripe_Working!$E:$E,Stripe_Working!$H:$H,A{r},Stripe_Working!$K:$K,"Valid"))',
     "I":G+'IF(C{r}="Closed Won",E{r}-H{r},0))',
     "J":G+'IF(C{r}<>"Closed Won",IF(H{r}>0,"Billed but not won","Not Won"),IF(H{r}=0,"Unbilled",IF(I{r}>0.005,"Underbilled",IF(I{r}<-0.005,"Overbilled","Matched")))))',
     "K":G+'F{r}-D{r}+1)',
     "L":G+'IF(C{r}="Closed Won",MAX(0,MIN(F{r},Summary_Dashboard!$B$2)-D{r}+1)/K{r}*E{r},0))',
     "M":G+'IF(C{r}="Closed Won",H{r}-L{r},0))',
     "N":G+'IF(C{r}="Closed Won",COUNTIF(C$2:C{r},"Closed Won"),""))'}
    hfmt={"H":CUR2,"I":CUR2,"K":"0","L":CUR2,"M":CUR2,"N":"0"}
    for r in range(2,HLAST+1):
        for c,f in hform.items():
            x=HW[f"{c}{r}"]; x.value=f.replace("{r}",str(r)); x.font=fn(); x.fill=fill(HELP); x.border=BOX
            if c in hfmt: x.number_format=hfmt[c]
    widths(HW,{"A":10,"B":26,"C":13,"D":12,"E":15,"F":12,"G":28,"H":15,"I":15,"J":19,"K":10,"L":17,"M":17,"N":9}); HW.freeze_panes="C2"
    for ws in (HW,SW):
        ws.conditional_formatting.add("J2:J500" if ws is HW else "J2:J500",FormulaRule(formula=['OR(J2="Duplicate",J2="Unbilled",J2="Underbilled",J2="Overbilled",J2="Billed but not won")'],fill=PatternFill("solid",bgColor="F8D7D3",fgColor="F8D7D3")))
    SW.conditional_formatting.add("R2:R500",FormulaRule(formula=['AND(R2<>"OK",R2<>"Excluded",R2<>"")'],fill=PatternFill("solid",bgColor="FFE8B5",fgColor="FFE8B5")))

    GRN=PatternFill("solid",bgColor="D9EEDF",fgColor="D9EEDF"); RED=PatternFill("solid",bgColor="F8D7D3",fgColor="F8D7D3")
    SW.conditional_formatting.add("N2:O500",FormulaRule(formula=['N2="Paid"'],fill=GRN))
    SW.conditional_formatting.add("O2:O500",FormulaRule(formula=['OR(O2="Past due",O2="Partially paid",O2="Duplicate")'],fill=RED))
    SW.conditional_formatting.add("J2:J500",FormulaRule(formula=['J2="Duplicate"'],fill=RED))
    # ---------------- RevRec_Waterfall ----------------
    MC=[chr(ord("H")+i) for i in range(12)]   # H..S
    head(RR,1,1,["Deal ID","Customer","Start Date","End Date","Total_Contract Value","Contract Days","Revenue Before Window"],dark=NAVY)
    for i,c in enumerate(MC):
        f="=Summary_Dashboard!$B$3" if i==0 else f"=EDATE({MC[i-1]}1,1)"
        x=RR[f"{c}1"]; x.value=f; x.number_format=MON; x.font=fn(bold=True,color=WHT); x.fill=fill(TEAL); x.alignment=Alignment(horizontal="center",vertical="center"); x.border=BOX
    head(RR,1,20,["Revenue After Window","Check"],dark=NAVY); RR.row_dimensions[1].height=32
    put(RR,"A2","TOTAL",bold=True,fl=LG)
    for c in "BCD": put(RR,f"{c}2",None,fl=LG)
    for c in ["E","G"]+MC+["T","U"]: put(RR,f"{c}2",f"=SUM({c}3:{c}1000)",CUR,bold=True,fl=LG)
    put(RR,"F2",None,fl=LG)
    for r in range(3,RLAST+1):
        g=f'=IF($A{r}="","",'
        rf={"A":'=IFERROR(INDEX(HubSpot_Working!$A:$A,MATCH(ROW()-2,HubSpot_Working!$N:$N,0)),"")',
            "B":g+'INDEX(HubSpot_Working!$B:$B,MATCH($A{r},HubSpot_Working!$A:$A,0)))',"C":g+'INDEX(HubSpot_Working!$D:$D,MATCH($A{r},HubSpot_Working!$A:$A,0)))',
            "D":g+'INDEX(HubSpot_Working!$F:$F,MATCH($A{r},HubSpot_Working!$A:$A,0)))',"E":g+'INDEX(HubSpot_Working!$E:$E,MATCH($A{r},HubSpot_Working!$A:$A,0)))',
            "F":g+'D{r}-C{r}+1)',"G":g+'MAX(0,MIN(D{r},$H$1-1)-C{r}+1)/F{r}*E{r})',
            "T":g+'MAX(0,D{r}-MAX(C{r},EOMONTH($S$1,0)+1)+1)/F{r}*E{r})',"U":g+'ROUND(E{r}-G{r}-SUM(H{r}:T{r}),2))'}
        for c in MC: rf[c]=g+f'IF(OR($C{r}>EOMONTH({c}$1,0),$D{r}<{c}$1),0,((MIN($D{r},EOMONTH({c}$1,0))-MAX($C{r},{c}$1)+1)/$F{r})*$E{r}))'
        for c,f in rf.items():
            x=RR[f"{c}{r}"]; x.value=f.replace("{r}",str(r)); x.font=fn(); x.border=BOX
            x.number_format = DT if c in "CD" else ("0" if c in "F" else (CUR if c not in "AB" else "General"))
            if c in MC or c in "GT": x.fill=fill(HELP) if c in "GT" else PatternFill()
    widths(RR,{"A":10,"B":26,"C":12,"D":12,"E":14,"F":10,"G":14,"T":14,"U":12,**{c:11 for c in MC}}); RR.freeze_panes="C3"

    # ---------------- Deferred_Rev ----------------
    DC=[chr(ord("B")+i) for i in range(12)]   # B..M
    RT=dict(zip(DC,MC))
    DR["A1"]="Deferred Revenue Rollforward"; DR["A1"].font=fn(size=16,bold=True,color=NAVY)
    DR["A2"]="Beginning + new valid billings - recognized revenue = ending deferred revenue. Months follow the schedule start date on Summary_Dashboard (B3)."; DR["A2"].font=fn(italic=True,color=GREY)
    put(DR,"A3","Item",bold=True,fl=NAVY,color=WHT)
    labs={4:"Beginning Deferred Revenue",5:"(+) New Billings (Stripe)",6:"(-) Revenue Recognized",7:"(=) Ending Deferred Revenue",9:"Check: billed to date minus revenue to date"}
    for r,t in labs.items(): put(DR,f"A{r}",t,bold=(r in (4,7)))
    for i,c in enumerate(DC):
        put(DR,f"{c}3",f"=RevRec_Waterfall!{RT[c]}1",MON,bold=True,fl=NAVY,color=WHT,align="center")
        put(DR,f"{c}4",'=SUMIFS(Stripe_Working!$E:$E,Stripe_Working!$K:$K,"Valid",Stripe_Working!$C:$C,"<"&B$3)-RevRec_Waterfall!$G$2' if i==0 else f"={DC[i-1]}7",CUR)
        put(DR,f"{c}5",f'=SUMIFS(Stripe_Working!$E:$E,Stripe_Working!$K:$K,"Valid",Stripe_Working!$C:$C,">="&{c}$3,Stripe_Working!$C:$C,"<="&EOMONTH({c}$3,0))',CUR)
        put(DR,f"{c}6",f"=RevRec_Waterfall!{RT[c]}2",CUR)
        put(DR,f"{c}7",f"={c}4+{c}5-{c}6",CUR,bold=True,fl=LG)
        put(DR,f"{c}9",f'=IF(ABS({c}7-(SUMIFS(Stripe_Working!$E:$E,Stripe_Working!$K:$K,"Valid",Stripe_Working!$C:$C,"<="&EOMONTH({c}$3,0))-(RevRec_Waterfall!$G$2+SUM($B6:{c}6))))<0.01,"Ties","Check")')
    okfmt(DR,"B9:M9","B9")
    DR["A11"]="A negative ending balance means revenue has been earned ahead of invoicing (unbilled AR / accrued revenue). A positive balance is deferred revenue (a liability)."; DR["A11"].font=fn(size=9,italic=True,color=GREY)
    widths(DR,{"A":50,**{c:12 for c in DC}}); DR.freeze_panes="B4"

    # ---------------- Summary_Dashboard ----------------
    SD.sheet_view.showGridLines=False
    widths(SD,{"A":44,"B":16,"C":38,"D":10,"E":3,"F":40,"G":12,"H":14})
    SD.merge_cells("A1:H1"); SD["A1"]="Order to Cash and Revenue Review"; SD["A1"].font=fn(size=18,bold=True,color=WHT); SD["A1"].alignment=Alignment(vertical="center",indent=1)
    for c in range(1,9): SD.cell(row=1,column=c).fill=fill(NAVY)
    SD.row_dimensions[1].height=36
    put(SD,"A2","As of Date",bold=True); put(SD,"B2",ASOF,DT,inp=True,fl=YEL,bold=True); SD["C2"]="Every aging and revenue formula reads this date."; SD["C2"].font=fn(italic=True,color=GREY,size=9)
    put(SD,"A3","Revenue Schedule Starts",bold=True); put(SD,"B3",first_month,DT,inp=True,fl=YEL,bold=True); SD["C3"]="First of 12 months. Uploaded data."; SD["C3"].font=fn(italic=True,color=GREY,size=9)
    H=lambda c: f"HubSpot_Working!${c}:${c}"; S=lambda c: f"Stripe_Working!${c}:${c}"
    sect(SD,"A5","A. Order to Cash"); head(SD,6,1,["Item","Amount","Notes","Check"],h=22)
    A=[("Total Contracted (Closed Won)",f'=SUMIFS({H("E")},{H("C")},"Closed Won")',"HubSpot contract value, Closed Won only"),
       ("(-) Unbilled or Underbilled Deals",f'=SUMIFS({H("I")},{H("I")},">0")',"Contract value minus billed"),
       ("(+) Overbilled vs Contract",f'=-SUMIFS({H("I")},{H("I")},"<0")',"Billed more than the contract"),
       ("(+) Billed, No Closed Won Deal",f'=B10_PLACEHOLDER',"Invoice has no deal, or deal not won"),
       ("(=) Valid Total Billed (Stripe)",f'=SUMIFS({S("E")},{S("K")},"Valid")',"No void, draft or duplicate"),
       ("(-) Total Cash Collected",f'=SUMIFS({S("F")},{S("K")},"Valid")',"Paid amount on valid invoices"),
       ("(=) Total Open AR",f'=SUM({S("L")})',"Balance on valid open invoices"),
       ("Memo: Duplicate Invoices Caught ($)",f'=SUMIFS({S("E")},{S("J")},"Duplicate")',"Same customer, amount, deal, month"),
       ("Memo: Duplicate Invoices Caught (#)",f'=COUNTIFS({S("J")},"Duplicate")',"")]
    for i,(l,f,n) in enumerate(A):
        r=7+i; put(SD,f"A{r}",l,bold=(l.startswith("(=)") or i==0)); put(SD,f"B{r}",f,NUM if "(#)" in l else CUR,bold=(l.startswith("(=)") or i==0)); put(SD,f"C{r}",n,italic=True,color=GREY); put(SD,f"D{r}",None)
    SD["B10"]=f'=B11-SUMIFS({H("H")},{H("C")},"Closed Won")'
    SD["D11"]='=IF(ABS(B7-B8+B9+B10-B11)<0.01,"Ties","Check")'; SD["D13"]='=IF(ABS(B11-B12-B13)<0.01,"Ties","Check")'
    for r in (11,13): SD[f"D{r}"].font=fn(bold=True); SD[f"D{r}"].alignment=Alignment(horizontal="center")
    okfmt(SD,"D11:D13","D11")
    sect(SD,"A17","B. AR Aging"); head(SD,18,1,["Bucket","Open AR ($)","Invoices (#)","% of total"],h=22)
    for i,b in enumerate(["Current","1 to 30","31 to 60","61 to 90","Over 90"]):
        r=19+i; put(SD,f"A{r}",b); put(SD,f"B{r}",f'=SUMIFS({S("L")},{S("N")},A{r})',CUR); put(SD,f"C{r}",f'=COUNTIFS({S("N")},A{r})',NUM,align="center"); put(SD,f"D{r}",f"=IF($B$24=0,0,B{r}/$B$24)",PCT)
    put(SD,"A24","Total Open AR",bold=True,fl=LG); put(SD,"B24","=SUM(B19:B23)",CUR,bold=True,fl=LG); put(SD,"C24","=SUM(C19:C23)",NUM,bold=True,fl=LG,align="left"); put(SD,"D24","=SUM(D19:D23)",PCT,bold=True,fl=LG)
    put(SD,"A25","Matches Open AR in Table A"); put(SD,"B25",'=IF(ABS(B24-B13)<0.01,"Ties","Check")',bold=True); okfmt(SD,"B25:B25","B25")
    # C. revenue
    sect(SD,"F5","C. Revenue and Deferred Revenue"); head(SD,6,6,["Item","Amount",""],h=22)
    mm='MATCH(DATE(YEAR($B$2),MONTH($B$2),1),Deferred_Rev!$B$3:$M$3,0)'
    Cc=[("Revenue This Month",f"=INDEX(Deferred_Rev!$B$6:$M$6,{mm})",CUR),
        ("Revenue Year to Date",'=SUMIFS(Deferred_Rev!$B$6:$M$6,Deferred_Rev!$B$3:$M$3,"<="&$B$2,Deferred_Rev!$B$3:$M$3,">="&DATE(YEAR($B$2),1,1))',CUR),
        ("Billings This Month",f"=INDEX(Deferred_Rev!$B$5:$M$5,{mm})",CUR),
        ("Deferred Revenue (Rollforward)",f"=INDEX(Deferred_Rev!$B$7:$M$7,{mm})",CUR),
        ("Deferred Revenue by Deal",f'=SUMIFS({H("M")},{H("M")},">0")',CUR),
        ("Unbilled AR by Deal",f'=-SUMIFS({H("M")},{H("M")},"<0")',CUR),
        ("Rollforward Matches Deals","=IF(ABS(G10-(G11-G12+B10))<0.01,\"Ties\",\"Check\")",None)]
    for i,(l,f,fm) in enumerate(Cc):
        r=7+i; put(SD,f"F{r}",l,bold=(i==3)); put(SD,f"G{r}",f,fm,bold=(i in (3,6)))
    SD["G13"].alignment=Alignment(horizontal="center"); okfmt(SD,"G13:G13","G13")
    sect(SD,"F15","D. Data Checks"); head(SD,16,6,["Check","Count","Amount ($)"],h=22)
    Dq=[("Duplicate Invoices",f'=COUNTIFS({S("J")},"Duplicate")',f'=SUMIFS({S("E")},{S("J")},"Duplicate")'),
        ("Void Invoices",f'=COUNTIFS({S("G")},"Void")',f'=SUMIFS({S("E")},{S("G")},"Void")'),
        ("Draft Invoices",f'=COUNTIFS({S("G")},"Draft")',f'=SUMIFS({S("E")},{S("G")},"Draft")'),
        ("Invoices Missing Deal ID",f'=COUNTIFS({S("R")},"Missing Deal ID")',f'=SUMIFS({S("E")},{S("R")},"Missing Deal ID")'),
        ("Invoices, No Matching Deal",f'=COUNTIFS({S("R")},"No matching deal")',f'=SUMIFS({S("E")},{S("R")},"No matching deal")'),
        ("Invoices on Deals Not Won",f'=COUNTIFS({S("R")},"Deal not won")',f'=SUMIFS({S("E")},{S("R")},"Deal not won")'),
        ("Closed Won, Not Invoiced",f'=COUNTIFS({H("J")},"Unbilled")',f'=SUMIFS({H("I")},{H("J")},"Unbilled")'),
        ("Closed Won, Underbilled",f'=COUNTIFS({H("J")},"Underbilled")',f'=SUMIFS({H("I")},{H("J")},"Underbilled")'),
        ("Closed Won, Overbilled",f'=COUNTIFS({H("J")},"Overbilled")',f'=-SUMIFS({H("I")},{H("J")},"Overbilled")')]
    for i,(l,c1,c2) in enumerate(Dq):
        r=17+i; put(SD,f"F{r}",l); put(SD,f"G{r}",c1,NUM); put(SD,f"H{r}",c2,CUR)
    # charts
    def bar(t,w,h):
        ch=BarChart(); ch.type="col"; ch.title=t; ch.width=w; ch.height=h; ch.legend=None; ch.roundedCorners=False; ch.gapWidth=60
        ch.x_axis.delete=False; ch.y_axis.delete=False; ch.y_axis.majorGridlines.spPr=GraphicalProperties(ln=None); ch.y_axis.majorGridlines.spPr.line.solidFill="E3E7EC"; return ch
    def colr(s,c): s.graphicalProperties.solidFill=c; s.graphicalProperties.line.solidFill=c
    c1=bar("Open AR by Aging",15.5,7.2); c1.add_data(Reference(SD,min_col=2,min_row=18,max_row=23),titles_from_data=True); c1.set_categories(Reference(SD,min_col=1,min_row=19,max_row=23))
    colr(c1.series[0],NAVY); c1.y_axis.number_format='$#,##0,"K"'; c1.dataLabels=DataLabelList(); c1.dataLabels.showVal=True; c1.dataLabels.showSerName=False; c1.dataLabels.showCatName=False; c1.dataLabels.showLegendKey=False; c1.dataLabels.numFmt='$#,##0,"K"'
    SD.add_chart(c1,"A28")
    c2=bar("Billings vs Revenue by Month",17.5,7.2)
    for rr in (5,6): c2.add_data(Reference(DR,min_col=1,max_col=13,min_row=rr),from_rows=True,titles_from_data=True)
    c2.set_categories(Reference(DR,min_col=2,max_col=13,min_row=3)); colr(c2.series[0],"B8C2CF"); colr(c2.series[1],TEAL); c2.y_axis.number_format='$#,##0,"K"'; c2.x_axis.number_format=MON
    c2.legend=Legend(); c2.legend.position="b"; SD.add_chart(c2,"F28")
    sect(SD,"A44","Executive Takeaways")
    pd_="SUMIFS(Stripe_Working!$L:$L,Stripe_Working!$M:$M,\">0\")"
    t1=('="1. Collect "&TEXT(B13,"$#,##0")&" in open AR. "&TEXT('+pd_+',"$#,##0")&" is past due and "&TEXT(B22+B23,"$#,##0")&" is over 60 days. Bill "&TEXT(H23+H24,"$#,##0")&" on "&(G23+G24)&" Closed Won deals not fully invoiced. "&G20&" invoice has no Deal ID, so check that one first."')
    t2=('="2. Void or refund "&G17&" duplicate invoices ("&TEXT(H17,"$#,##0")&"). Also look at "&TEXT(G21+G22,"0")&" invoices ("&TEXT(H21+H22,"$#,##0")&") with no matching deal or on a deal not won, and "&G25&" overbilled deal ("&TEXT(H25,"$#,##0")&")."')
    t3=('="3. Rule going forward: Stripe should not send an invoice unless the Deal ID exists in HubSpot, the deal is Closed Won, and the amount equals the Contract Value. That rule would catch "&TEXT(G17+G20+G21+G22+G25,"0")&" of the issues here."')
    for i,t in enumerate((t1,t2,t3)):
        r=45+i; SD.merge_cells(f"A{r}:H{r}"); SD[f"A{r}"]=t; SD[f"A{r}"].font=fn(); SD[f"A{r}"].alignment=Alignment(wrap_text=True,vertical="top"); SD.row_dimensions[r].height=34
    SD.page_setup.orientation="landscape"; SD.page_setup.fitToWidth=1; SD.page_setup.fitToHeight=0; SD.sheet_properties.pageSetUpPr.fitToPage=True

    # ---------------- Support + green Dashboard ----------------
    from openpyxl.chart import DoughnutChart
    from openpyxl.chart.series import DataPoint
    from openpyxl.worksheet.datavalidation import DataValidation
    GD="1B5E20"; GM="2E7D32"; GL="81C784"; GP="C8E6C9"; GBG="E8F5E9"
    SP=wb.create_sheet("Support"); DB=wb.create_sheet("Dashboard"); DB.sheet_view.showGridLines=False
    SWr=lambda c: f"Stripe_Working!${c}:${c}"; SPr=lambda c: f"Support!${c}:${c}"
    def sh(ws,ref,t):
        c=ws[ref]; c.value=t; c.font=fn(bold=True,color=WHT); c.fill=fill(GM); c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); c.border=BOX
    def sv(ref,val,f=None,bold=False):
        c=SP[ref]; c.value=val; c.font=fn(bold=bold); c.border=BOX
        if f: c.number_format=f
    SP["A1"]="Support: tables that feed the Dashboard (formulas, do not edit)"; SP["A1"].font=fn(bold=True,size=12)
    crit=[("Month start",'=IF(Dashboard!$S$6="All",DATE(2000,1,1),INDEX($U$11:$U$22,MATCH(Dashboard!$S$6,$T$11:$T$22,0)))'),("Month end (exclusive)",'=IF(Dashboard!$S$6="All",DATE(2100,1,1),EDATE(B2,1))'),
          ("Status",'=IF(Dashboard!$S$9="All","*",Dashboard!$S$9)'),("Aging bucket",'=IF(Dashboard!$S$12="All","*",Dashboard!$S$12)'),("Deal check",'=IF(Dashboard!$S$15="All","*",Dashboard!$S$15)')]
    for i,(l,f) in enumerate(crit): sv(f"A{2+i}",l,bold=True); sv(f"B{2+i}",f)
    SP["B2"].number_format=DT; SP["B3"].number_format=DT
    def cr(month=True,status=True,aging=True,check=True):
        p=[]
        if month: p+=[f'{SPr("AD")},">="&Support!$B$2',f'{SPr("AD")},"<"&Support!$B$3']
        if status: p.append(f'{SWr("O")},Support!$B$4')
        if aging: p.append(f'{SWr("N")},Support!$B$5')
        if check: p.append(f'{SWr("R")},Support!$B$6')
        return ",".join(p)
    VAL=f'{SWr("K")},"Valid"'
    kp=[("# of Invoices",f'=SUM({SPr("AE")})',"#,##0"),("Total Billed",f'=SUMIFS({SWr("E")},{VAL},{cr()})',"$#,##0"),("Total Collected",f'=SUMIFS({SWr("F")},{VAL},{cr()})',"$#,##0"),
        ("Open AR",f'=SUMIFS({SWr("L")},{cr()})',"$#,##0"),("Past Due AR",f'=SUMIFS({SWr("L")},{SWr("M")},">0",{cr()})',"$#,##0")]
    for i,(l,f,fm) in enumerate(kp): sv(f"G{2+i}",l,bold=True); sv(f"H{2+i}",f,fm)
    # lists
    for c,t in (("T","Month"),("U","Month start"),("W","Status"),("Y","Aging bucket"),("AA","Deal check")): sh(SP,f"{c}9",t)
    for c in ("T","W","Y","AA"): sv(f"{c}10","All")
    for k in range(12):
        sv(f"U{11+k}","=Summary_Dashboard!$B$3" if k==0 else f"=EDATE(U{10+k},1)",DT); sv(f"T{11+k}",f'=TEXT(U{11+k},"mmm yy")')
    STAT=["Paid","Partially paid","Past due","Open","Duplicate","Void","Draft"]; AGE=["Current","1 to 30","31 to 60","61 to 90","Over 90"]; CHK=["OK","Missing Deal ID","No matching deal","Deal not won"]
    for k,v in enumerate(STAT): sv(f"W{11+k}",v)
    for k,v in enumerate(["Paid"]+AGE+["Excluded"]): sv(f"Y{11+k}",v)
    for k,v in enumerate(CHK+["Excluded"]): sv(f"AA{11+k}",v)
    # table 1 aging
    sh(SP,"A9","Aging Bucket"); sh(SP,"B9","Open AR")
    for k,b in enumerate(AGE): sv(f"A{10+k}",b); sv(f"B{10+k}",f'=SUMIFS({SWr("L")},{SWr("N")},A{10+k},{cr(aging=False)})',"$#,##0")
    # table 2 status
    sh(SP,"D9","Status"); sh(SP,"E9","Number of Invoices")
    for k,v in enumerate(STAT): sv(f"D{10+k}",v); sv(f"E{10+k}",f'=COUNTIFS({SWr("O")},D{10+k},{cr(status=False)})',"0")
    # table 3 collection rate
    sh(SP,"G9","Values"); sh(SP,"H9","Share")
    sv("G10","Collected"); sv("H10",f'=IFERROR(SUMIFS({SWr("F")},{VAL},{cr()})/SUMIFS({SWr("E")},{VAL},{cr()}),0)',"0.0%"); sv("G11","Still open"); sv("H11","=1-H10","0.0%")
    # table 4 billed by deal check
    sh(SP,"J9","Deal Check"); sh(SP,"K9","Billed ($)")
    for k,v in enumerate(CHK): sv(f"J{10+k}",v); sv(f"K{10+k}",f'=SUMIFS({SWr("E")},{SWr("R")},J{10+k},{cr(check=False)})',"$#,##0")
    # table 5 by month
    sh(SP,"M9","Month"); sh(SP,"N9","Billed"); sh(SP,"O9","Collected")
    for k in range(12):
        mo=f'{SPr("AD")},">="&$U${11+k},{SPr("AD")},"<"&EDATE($U${11+k},1)'
        sv(f"M{10+k}",f"=T{11+k}"); sv(f"N{10+k}",f'=SUMIFS({SWr("E")},{VAL},{mo},{cr(month=False)})',"$#,##0"); sv(f"O{10+k}",f'=SUMIFS({SWr("F")},{VAL},{mo},{cr(month=False)})',"$#,##0")
    # table 6 past due share
    sh(SP,"Q9","Values"); sh(SP,"R9","Share")
    sv("Q10","Past due"); sv("R10",f'=IFERROR(SUMIFS({SWr("L")},{SWr("M")},">0",{cr()})/SUMIFS({SWr("L")},{cr()}),0)',"0.0%"); sv("Q11","Not past due"); sv("R11","=1-R10","0.0%")
    # invoice list
    SP["A18"]="Invoices matching the filters"; SP["A18"].font=fn(bold=True)
    for c,t in zip("ABCDEFGH",["#","Invoice","Customer","Deal ID","Due date","Amount","Open balance","Status"]): sh(SP,f"{c}19",t)
    for n in range(1,13):
        r=19+n; sv(f"A{r}",n)
        for c,sc,fm in (("B","A",None),("C","B",None),("D","H",None),("E","D",DT),("F","E","$#,##0"),("G","L","$#,##0"),("H","O",None)):
            sv(f"{c}{r}",f'=IFERROR(INDEX({SWr(sc)},MATCH($A{r},{SPr("AF")},0)),"")',fm)
    # helper block per invoice row
    for c,t in zip(["AD","AE","AF"],["Invoice Date (clean)","In Dashboard Filter","Filter Match #"]): sh(SP,f"{c}1",t)
    for r in range(2,SLAST+1):
        g=f'=IF(Stripe_Working!$A{r}="","",'
        SP[f"AD{r}"]=g+f'IFERROR(Stripe_Working!C{r}+0,""))'
        SP[f"AE{r}"]=g+f'IF(AND(AD{r}>=$B$2,AD{r}<$B$3,OR($B$4="*",Stripe_Working!O{r}=$B$4),OR($B$5="*",Stripe_Working!N{r}=$B$5),OR($B$6="*",Stripe_Working!R{r}=$B$6)),1,0))'
        SP[f"AF{r}"]=g+f'IF(AE{r}=1,COUNTIF(AE$2:AE{r},1),""))'
        for c in ("AD","AE","AF"): SP[f"{c}{r}"].font=fn(); SP[f"{c}{r}"].number_format="0"
        SP[f"AD{r}"].number_format=DT
    for c,w in {"A":22,"B":18,"C":26,"D":16,"E":16,"F":12,"G":16,"H":14,"I":3,"J":18,"K":14,"L":3,"M":12,"N":14,"O":14,"P":3,"Q":14,"R":10,"S":3,"T":12,"U":12,"V":3,"W":15,"X":3,"Y":15,"Z":3,"AA":18,"AB":3,"AC":3,"AD":14,"AE":12,"AF":12}.items(): SP.column_dimensions[c].width=w
    SP.row_dimensions[1].height=30
    # ---- Dashboard ----
    DB.column_dimensions["A"].width=2
    for c in "BCDEFGHIJKLMNOP": DB.column_dimensions[c].width=9.5
    DB.column_dimensions["Q"].width=2; DB.column_dimensions["R"].width=2
    for c in "STU": DB.column_dimensions[c].width=11
    for r in range(1,62):
        for c in range(2,22): DB.cell(row=r,column=c).fill=fill(GBG)
    DB.merge_cells("B1:U3"); DB["B1"]="BILLING DASHBOARD"; DB["B1"].font=fn(size=22,bold=True,color=WHT); DB["B1"].alignment=Alignment(horizontal="center",vertical="center")
    for r in (1,2,3):
        for c in range(2,22): DB.cell(row=r,column=c).fill=fill(GD)
    DB["B4"]='="As of "&TEXT(Summary_Dashboard!B2,"mmm d, yyyy")&"   |   Uploaded data   |   All numbers come from the Support tab"'; DB["B4"].font=fn(size=9,italic=True,color=GM)
    crd=[("# OF INVOICES","H2","#,##0"),("TOTAL BILLED","H3","$#,##0"),("TOTAL COLLECTED","H4","$#,##0"),("OPEN AR","H5","$#,##0"),("PAST DUE AR","H6","$#,##0")]
    GBOX=Border(left=Side(style="medium",color=GM),right=Side(style="medium",color=GM),top=Side(style="medium",color=GM),bottom=Side(style="medium",color=GM))
    for i,(l,ref,fm) in enumerate(crd):
        c=chr(ord("B")+3*i); e=chr(ord(c)+2)
        DB.merge_cells(f"{c}6:{e}7"); DB.merge_cells(f"{c}8:{e}8")
        DB[f"{c}6"]=f"=Support!{ref}"; DB[f"{c}6"].font=fn(size=22,bold=True,color=GD); DB[f"{c}6"].number_format=fm; DB[f"{c}6"].alignment=Alignment(horizontal="center",vertical="center")
        DB[f"{c}8"]=l; DB[f"{c}8"].font=fn(size=9,bold=True,color=GM); DB[f"{c}8"].alignment=Alignment(horizontal="center")
        for r in (6,7,8):
            for cc in range(ord(c),ord(e)+1):
                x=DB[f"{chr(cc)}{r}"]; x.fill=fill(WHT)
                x.border=Border(left=Side(style="medium",color=GM) if cc==ord(c) else None,right=Side(style="medium",color=GM) if cc==ord(e) else None,top=Side(style="medium",color=GM) if r==6 else None,bottom=Side(style="medium",color=GM) if r==8 else None)
    for hr,vr,t in [(5,6,"MONTH"),(8,9,"STATUS"),(11,12,"AGING BUCKET"),(14,15,"DEAL CHECK")]:
        DB.merge_cells(f"S{hr}:U{hr}"); DB.merge_cells(f"S{vr}:U{vr}")
        DB[f"S{hr}"]=t+"  ▼"; DB[f"S{hr}"].font=fn(size=10,bold=True,color=WHT); DB[f"S{hr}"].alignment=Alignment(horizontal="center",vertical="center")
        for c in "STU": DB[f"{c}{hr}"].fill=fill(GM)
        DB[f"S{vr}"]="All"; DB[f"S{vr}"].font=fn(size=12,bold=True,color=GD); DB[f"S{vr}"].alignment=Alignment(horizontal="center",vertical="center")
        for c in "STU": DB[f"{c}{vr}"].fill=fill(WHT); DB[f"{c}{vr}"].border=GBOX
    for r in (6,9,12,15): DB.row_dimensions[r].height=24
    DB.merge_cells("S17:U19"); DB["S17"]="Tap a white box, pick a value from the list, set back to All to clear."; DB["S17"].font=fn(size=8,italic=True,color=GM); DB["S17"].alignment=Alignment(wrap_text=True,vertical="top")
    for ref,lst in (("S6","$T$10:$T$22"),("S9","$W$10:$W$17"),("S12","$Y$10:$Y$16"),("S15","$AA$10:$AA$15")):
        dv=DataValidation(type="list",formula1="=Support!"+lst,allow_blank=False); DB.add_data_validation(dv); dv.add(ref)
    DB.merge_cells("S21:U21"); DB["S21"]="OPEN AR PAST DUE"; DB["S21"].font=fn(size=10,bold=True,color=WHT); DB["S21"].alignment=Alignment(horizontal="center")
    for c in "STU": DB[f"{c}21"].fill=fill(GM)
    DB.merge_cells("S22:U22"); DB["S22"]='=REPT("█",ROUND(Support!R10*12,0))&REPT("░",12-ROUND(Support!R10*12,0))'; DB["S22"].font=fn(size=11,bold=True,color=GM); DB["S22"].alignment=Alignment(horizontal="center")
    DB.merge_cells("S23:U24"); DB["S23"]="=Support!R10"; DB["S23"].number_format="0.0%"; DB["S23"].font=fn(size=24,bold=True,color=GD); DB["S23"].alignment=Alignment(horizontal="center",vertical="center")
    for r in (22,23,24):
        for c in "STU": DB[f"{c}{r}"].fill=fill(WHT)
    def gstyle(ch,t,w,h):
        ch.title=t; ch.width=w; ch.height=h; ch.legend=None; ch.roundedCorners=False; ch.x_axis.delete=False; ch.y_axis.delete=False
        ch.y_axis.majorGridlines.spPr=GraphicalProperties(ln=None); ch.y_axis.majorGridlines.spPr.line.solidFill="D5E8D4"
    def gcol(s,c): s.graphicalProperties.solidFill=c; s.graphicalProperties.line.solidFill=c
    def labels(ch,fmt='$#,##0'):
        ch.dataLabels=DataLabelList(); x=ch.dataLabels; x.showVal=True; x.showCatName=False; x.showSerName=False; x.showLegendKey=False; x.showPercent=False; x.numFmt=fmt
    def pts(ser,cols):
        for i,cl in enumerate(cols):
            dp=DataPoint(idx=i); dp.graphicalProperties.solidFill=cl; dp.graphicalProperties.line.solidFill="FFFFFF"; ser.dPt.append(dp)
    def legend(ch): ch.legend=Legend(); ch.legend.position="b"; ch.legend.overlay=False
    g1=BarChart(); g1.type="col"; gstyle(g1,"Open AR by Aging Bucket",9.4,6.8)
    g1.add_data(Reference(SP,min_col=2,min_row=9,max_row=14),titles_from_data=True); g1.set_categories(Reference(SP,min_col=1,min_row=10,max_row=14))
    gcol(g1.series[0],GM); pts(g1.series[0],["1B5E20","388E3C","66BB6A","A5D6A7","C8E6C9"]); labels(g1); g1.gapWidth=50; g1.y_axis.number_format='$#,##0,"K"'; DB.add_chart(g1,"B10")
    g2=DoughnutChart(); g2.title="Number of Invoices by Status"; g2.width=9.4; g2.height=6.8; g2.holeSize=60; g2.roundedCorners=False
    g2.add_data(Reference(SP,min_col=5,min_row=9,max_row=16),titles_from_data=True); g2.set_categories(Reference(SP,min_col=4,min_row=10,max_row=16))
    pts(g2.series[0],["2E7D32","F9A825","D32F2F","A5D6A7","8D6E63","9E9E9E","BDBDBD"]); labels(g2,"0"); legend(g2); DB.add_chart(g2,"G10")
    g3=BarChart(); g3.type="col"; gstyle(g3,"Billed by Deal Check",9.4,6.8)
    g3.add_data(Reference(SP,min_col=11,min_row=9,max_row=13),titles_from_data=True); g3.set_categories(Reference(SP,min_col=10,min_row=10,max_row=13))
    gcol(g3.series[0],GM); pts(g3.series[0],["2E7D32","F9A825","D32F2F","8D6E63"]); labels(g3); g3.gapWidth=50; g3.y_axis.number_format='$#,##0,"K"'; DB.add_chart(g3,"L10")
    g4=DoughnutChart(); g4.title="Collection Rate (%)"; g4.width=9.4; g4.height=6.8; g4.holeSize=65; g4.roundedCorners=False
    g4.add_data(Reference(SP,min_col=8,min_row=9,max_row=11),titles_from_data=True); g4.set_categories(Reference(SP,min_col=7,min_row=10,max_row=11))
    pts(g4.series[0],[GM,"E0E0E0"]); labels(g4,"0%"); g4.dataLabels.showCatName=False; DB.add_chart(g4,"B25")
    g5=BarChart(); g5.type="col"; gstyle(g5,"Billed vs Collected by Month",19.2,6.8)
    g5.add_data(Reference(SP,min_col=14,max_col=15,min_row=9,max_row=21),titles_from_data=True); g5.set_categories(Reference(SP,min_col=13,min_row=10,max_row=21))
    gcol(g5.series[0],GD); gcol(g5.series[1],GL); g5.gapWidth=60; g5.y_axis.number_format='$#,##0,"K"'; legend(g5); DB.add_chart(g5,"G25")
    DB["B40"]="INVOICES MATCHING THE FILTERS"; DB["B40"].font=fn(size=11,bold=True,color=GD)
    DB["G40"]=f'="Showing "&MIN(12,SUM({SPr("AE")}))&" of "&SUM({SPr("AE")})&" matching invoices"'; DB["G40"].font=fn(size=9,italic=True,color=GM)
    heads=[("B","#"),("C","Invoice"),("D","Customer"),("F","Deal ID"),("H","Due date"),("I","Amount"),("J","Open"),("K","Status")]
    for c,t in heads: sh(DB,f"{c}41",t)
    for a,b in (("D","E"),("F","G"),("K","L")): DB.merge_cells(f"{a}41:{b}41"); DB[f"{b}41"].fill=fill(GM)
    smap={"B":"A","C":"B","D":"C","F":"D","H":"E","I":"F","J":"G","K":"H"}
    for n in range(1,13):
        r=41+n
        for a,b in (("D","E"),("F","G"),("K","L")): DB.merge_cells(f"{a}{r}:{b}{r}")
        for c,sc in smap.items(): DB[f"{c}{r}"]=f"=Support!{sc}{19+n}"
        DB[f"H{r}"].number_format="mm/dd/yy"; DB[f"I{r}"].number_format='$#,##0'; DB[f"J{r}"].number_format='$#,##0'
        for c in "BCDEFGHIJKL":
            x=DB[f"{c}{r}"]; x.font=fn(size=10); x.fill=fill(WHT if n%2 else GP); x.border=Border(bottom=Side(style="thin",color=GL))
    DB.page_setup.orientation="landscape"; DB.page_setup.fitToWidth=1; DB.page_setup.fitToHeight=0; DB.sheet_properties.pageSetUpPr.fitToPage=True; DB.print_area="A1:V55"
    SP.sheet_properties.tabColor=GM; DB.sheet_properties.tabColor=GD

    # ---------------- Read_Me ----------------
    RM["A1"]="Read Me"; RM["A1"].font=fn(size=16,bold=True,color=NAVY); RM["A2"]="Order to cash and revenue. Uploaded data."; RM["A2"].font=fn(italic=True,color=GREY)
    lines=[("TABS",None),("Summary_Dashboard","Summary of everything. As of date in B2, revenue schedule start in B3."),("HubSpot_Working","Deals (blue text) with helper columns H to N."),
    ("Stripe_Working","Invoices (blue text) with helper columns I to R."),("RevRec_Waterfall","One row per Closed Won deal, revenue split by month."),("Deferred_Rev","Monthly deferred revenue rollforward."),("Dashboard, Support","Green dashboard with dropdown filters (Month, Status, Aging Bucket, Deal Check). Support holds the tables it reads."),
    ("Raw_HubSpot, Raw_Stripe","Original exports, not changed."),
    ("HOW TO REFRESH",None),("1","Paste new data as values into the blue text columns on the two Working tabs, starting at row 2."),
    ("2","Change the as of date (use a month end) and the revenue start month on Summary_Dashboard."),("3","Formulas are filled to row 200 (HubSpot), 400 (Stripe) and 202 (RevRec). For more rows, drag the last row down."),
    ("4","Check that every Ties cell says Ties."),
    ("HOW IT WORKS",None),("Valid Flag","Void, Draft and Duplicate invoices are excluded. Billings, cash and AR use Valid only."),("Status","Calculated from the balance: Paid, Partially paid, Past due or Open. Void, Draft and Duplicate are shown as is. Stripe Status in column G is the original export value."),
    ("Duplicate Check","The second invoice with the same customer, amount, deal and month is marked Duplicate."),
    ("Billing Status","Closed Won deals show Unbilled, Underbilled, Overbilled or Matched."),
    ("Revenue","Subscriptions are recognized by day over the contract. Onboarding fees are recognized in the month they start. Confirm with ASC 606 for real contracts."),
    ("Deferred vs Unbilled","Billed minus recognized. Positive is deferred revenue, negative is unbilled AR."),
    ("Month End","Use a month end as the as of date so revenue ties to the monthly rollforward."),
    ("NOTE",None),("Source","Built by the SaaS Billing & Revenue Hub app from the uploaded Raw_HubSpot and Raw_Stripe sheets. Open in Excel or Google Sheets so the formulas calculate.")]
    r=4
    for a,b in lines:
        if b is None: RM[f"A{r}"]=a; RM[f"A{r}"].font=fn(bold=True,size=11,color=NAVY); RM[f"A{r}"].fill=fill(SOFT); RM[f"B{r}"].fill=fill(SOFT)
        else: RM[f"A{r}"]=a; RM[f"B{r}"]=b; RM[f"A{r}"].font=fn(bold=True); RM[f"B{r}"].font=fn(); RM[f"B{r}"].alignment=Alignment(wrap_text=True,vertical="top")
        r+=1
    widths(RM,{"A":28,"B":120})
    for ws,c in ((SD,NAVY),(HW,TEAL),(SW,TEAL),(RR,"B58900"),(DR,"B58900"),(RM,GREY),(RH,GREY),(RS,GREY)): ws.sheet_properties.tabColor=c
    wb._sheets=[wb[n] for n in ["Summary_Dashboard","Dashboard","HubSpot_Working","Stripe_Working","RevRec_Waterfall","Deferred_Rev","Support","Read_Me","Raw_HubSpot","Raw_Stripe"]]
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

