from typing import Dict, List, Optional
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.utils import get_column_letter

NAVY="1F2A44"; ACCENT="2E6F9E"; ACCENT_LIGHT="DCE9F2"; GREEN="3D8B58"; RED="B23B3B"; LIGHT_GREY="F4F5F7"
FN="Calibri"
TF=Font(name=FN,size=16,bold=True,color="FFFFFF"); STF=Font(name=FN,size=10,italic=True,color="FFFFFF")
HF=Font(name=FN,size=10,bold=True,color="FFFFFF"); LF=Font(name=FN,size=10,bold=True,color=NAVY)
CF=Font(name=FN,size=10,color="333333"); MF=Font(name=FN,size=10,bold=True,color=NAVY)
TFill=PatternFill("solid",fgColor=NAVY); HFill=PatternFill("solid",fgColor=ACCENT)
MFill=PatternFill("solid",fgColor=ACCENT_LIGHT); BFill=PatternFill("solid",fgColor=LIGHT_GREY)
THIN=Side(style="thin",color="D0D3D8"); BDR=Border(left=THIN,right=THIN,top=THIN,bottom=THIN)
PCT={"Win rate (kumul.)"}

def _title(ws, title, subtitle, last_col):
    ws.merge_cells(f"A1:{last_col}1"); ws.merge_cells(f"A2:{last_col}2")
    ws["A1"]=title; ws["A1"].font=TF; ws["A1"].fill=TFill; ws["A1"].alignment=Alignment(horizontal="left",vertical="center",indent=1)
    ws["A2"]=subtitle; ws["A2"].font=STF; ws["A2"].fill=TFill; ws["A2"].alignment=Alignment(horizontal="left",vertical="center",indent=1)
    ws.row_dimensions[1].height=28; ws.row_dimensions[2].height=18

def write_table(ws, start_row, week_labels_display, stage_order, stage_data, metrics_order, metrics):
    n = len(week_labels_display); r = start_row; header_row = r
    ws.cell(r,1,"Stage / Metrika").font=HF; ws.cell(r,1).fill=HFill; ws.cell(r,1).border=BDR
    for w,lbl in enumerate(week_labels_display):
        c=ws.cell(r,2+w,lbl); c.font=HF; c.fill=HFill; c.alignment=Alignment(horizontal="center"); c.border=BDR
    r+=1; stage_rows={}
    for i,stage in enumerate(stage_order):
        band=BFill if i%2==0 else PatternFill(fill_type=None)
        ws.cell(r,1,stage).font=CF; ws.cell(r,1).fill=band; ws.cell(r,1).border=BDR
        vals=stage_data.get(stage,[0]*n)
        for w in range(n):
            c=ws.cell(r,2+w,vals[w]); c.number_format="#,##0"; c.font=CF; c.fill=band; c.border=BDR; c.alignment=Alignment(horizontal="right")
        stage_rows[stage]=r; r+=1
    r+=1; metric_rows={}
    for name in metrics_order:
        vals=metrics.get(name)
        if vals is None: continue
        ws.cell(r,1,name).font=MF; ws.cell(r,1).fill=MFill; ws.cell(r,1).border=BDR
        for w in range(n):
            c=ws.cell(r,2+w,vals[w]); c.font=MF; c.fill=MFill; c.border=BDR; c.alignment=Alignment(horizontal="right")
            c.number_format="0.0%" if name in PCT else "#,##0"
        metric_rows[name]=r; r+=1
    for ci in range(1,2+n): ws.column_dimensions[get_column_letter(ci)].width=13 if ci>1 else 30
    return header_row, stage_rows, metric_rows, r

def _add_charts(ws, anchor_row, header_row, stage_rows, metric_rows, has_goal, n_weeks, week_dates_for_title=None):
    mc = 1+n_weeks
    cats = Reference(ws,min_col=2,max_col=mc,min_row=header_row,max_row=header_row)
    def line(title, labels, y="Kč"):
        ch=LineChart(); ch.title=title; ch.style=10; ch.y_axis.title=y; ch.height=7.5; ch.width=15
        for lbl in labels:
            if lbl not in metric_rows: continue
            d=Reference(ws,min_col=1,max_col=mc,min_row=metric_rows[lbl],max_row=metric_rows[lbl])
            ch.add_data(d,titles_from_data=True,from_rows=True)
        ch.set_categories(cats); return ch
    def bar(title, lbl, color=ACCENT, y="Kč"):
        ch=BarChart(); ch.type="col"; ch.title=title; ch.style=10; ch.y_axis.title=y; ch.height=7.5; ch.width=15
        d=Reference(ws,min_col=1,max_col=mc,min_row=metric_rows[lbl],max_row=metric_rows[lbl])
        ch.add_data(d,titles_from_data=True,from_rows=True); ch.set_categories(cats)
        ch.series[0].graphicalProperties.solidFill=color; return ch
    def funnel():
        ch=BarChart(); ch.type="bar"
        sfx=f" ({week_dates_for_title[-1]})" if week_dates_for_title else ""
        ch.title=f"Funnel – rozpad pipeline{sfx}"; ch.style=10; ch.x_axis.title="Kč"; ch.height=7.8; ch.width=15
        lc=1+n_weeks
        d=Reference(ws,min_col=lc,max_col=lc,min_row=min(stage_rows.values()),max_row=max(stage_rows.values()))
        cs=Reference(ws,min_col=1,max_col=1,min_row=min(stage_rows.values()),max_row=max(stage_rows.values()))
        ch.add_data(d,titles_from_data=False); ch.set_categories(cs)
        ch.series[0].graphicalProperties.solidFill=GREEN; ch.legend=None; return ch
    charts=[bar("Týdenní změna pipeline","Changes in pipeline"),
            line("Won vs. Lost (týdně)",["Won","Lost"]),
            funnel()]
    if has_goal: charts.append(line("Tempo k cíli – Won vs. Goal",["Won (kumulativně)","Goal (kumulativně)"]))
    charts.append(line("Win rate v čase",["Win rate (kumul.)"],y="%"))
    charts.append(line("Průměrná velikost dealu",["Prům. velikost dealu"]))
    r=anchor_row; cols=["A","H"]
    for i,ch in enumerate(charts):
        ws.add_chart(ch,f"{cols[i%2]}{r}")
        if i%2==1: r+=16
    return anchor_row+((len(charts)+1)//2)*16

METRICS_ORDER=["Won","Lost","Pipeline till end of year","Changes in pipeline","Rolling 18",
               "Win rate (kumul.)","Prům. velikost dealu","Won (kumulativně)","Goal (kumulativně)"]

def build_person_sheet(wb, owner, sheet_data, stage_order, week_labels_display, week_dates_for_title):
    ws=wb.create_sheet(owner[:31]); n=len(week_labels_display); lc=get_column_letter(1+n)
    goal=sheet_data.get("annual_goal")
    goal_txt=f"Roční cíl: {goal:,.0f} Kč".replace(",","_").replace("_"," ") if goal else "Roční cíl: zatím nestanoven"
    _title(ws,f"{owner} — Sales Pipeline Report",f"Týdenní přehled pipeline, Kč  ·  {goal_txt}",lc)
    hr,sr,mr,er=write_table(ws,4,week_labels_display,stage_order,sheet_data["stage_weighted"],METRICS_ORDER,sheet_data["metrics"])
    _add_charts(ws,er+2,hr,sr,mr,sheet_data["metrics"].get("Goal (kumulativně)") is not None,n,week_dates_for_title)
    ws.freeze_panes="B5"; ws.sheet_view.showGridLines=False

def build_aggregation_sheet(wb, agg_data, stage_order, week_labels_display, week_dates_for_title, leaderboard):
    ws=wb.create_sheet("Aggregation"); n=len(week_labels_display); lc=get_column_letter(1+n)
    _title(ws,"Tapix — Agregovaný Sales Pipeline Report","Všichni obchodníci součtem, Kč",lc)
    hr,sr,mr,er=write_table(ws,4,week_labels_display,stage_order,agg_data["stage_weighted"],METRICS_ORDER,agg_data["metrics"])
    end=_add_charts(ws,er+2,hr,sr,mr,agg_data["metrics"].get("Goal (kumulativně)") is not None,n,week_dates_for_title)
    lb=end+2; ws.cell(lb,1,"Žebříček obchodníků (Won celkem)").font=LF; lb+=1
    ws.cell(lb,1,"Obchodník").font=HF; ws.cell(lb,1).fill=HFill
    ws.cell(lb,2,"Won celkem (Kč)").font=HF; ws.cell(lb,2).fill=HFill
    lbh=lb; lb+=1; lbf=lb
    for name,total in leaderboard:
        ws.cell(lb,1,name).font=CF
        c=ws.cell(lb,2,round(total)); c.number_format="#,##0"; c.font=CF; lb+=1
    lbl=max(lb-1,lbf)
    ch=BarChart(); ch.type="bar"; ch.title="Žebříček obchodníků podle celkového Won"; ch.style=10; ch.height=8; ch.width=15
    d=Reference(ws,min_col=2,max_col=2,min_row=lbh,max_row=lbl)
    cs=Reference(ws,min_col=1,max_col=1,min_row=lbf,max_row=lbl)
    ch.add_data(d,titles_from_data=True); ch.set_categories(cs)
    ch.series[0].graphicalProperties.solidFill=ACCENT; ch.legend=None
    ws.add_chart(ch,f"D{lbh}"); ws.freeze_panes="B5"; ws.sheet_view.showGridLines=False
    ws.column_dimensions["A"].width=30

def build_ledger_sheet(wb, title, rows, color):
    ws=wb.create_sheet(title)
    _title(ws,f"{title} Deals — automatický ledger","Generováno týdně z HubSpotu","F")
    headers=["Deal ID","Deal Name","Company","Deal Owner","Close Date","Amount (Kč)","Weeknum"]
    r=4
    for ci,h in enumerate(headers):
        c=ws.cell(r,1+ci,h); c.font=HF; c.fill=PatternFill("solid",fgColor=color); c.border=BDR
    r+=1
    for row in sorted(rows, key=lambda x: x[4] if x[4] else date.min):
        for ci,val in enumerate(row):
            c=ws.cell(r,1+ci,val); c.font=CF; c.border=BDR
            if ci==4: c.number_format="dd.mm.yyyy"
            if ci==5: c.number_format="#,##0"
        r+=1
    for ci,w in enumerate([10,26,16,14,13,14,9]): ws.column_dimensions[get_column_letter(1+ci)].width=w
    ws.freeze_panes="A5"; ws.sheet_view.showGridLines=False

def new_workbook():
    wb=Workbook(); wb.remove(wb.active); return wb
