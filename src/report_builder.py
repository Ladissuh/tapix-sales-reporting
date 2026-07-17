from typing import Dict, List, Optional
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import RichTextProperties, Paragraph, ParagraphProperties, CharacterProperties
from openpyxl.utils import get_column_letter

NAVY="1F2A44"; ACCENT="2E6F9E"; ACCENT_LIGHT="DCE9F2"; GREEN="3D8B58"; RED="B23B3B"; LIGHT_GREY="F4F5F7"
GOLD="C9A227"; PLUM="6B4C7A"
FN="Calibri"
TF=Font(name=FN,size=16,bold=True,color="FFFFFF"); STF=Font(name=FN,size=10,italic=True,color="FFFFFF")
HF=Font(name=FN,size=10,bold=True,color="FFFFFF"); LF=Font(name=FN,size=10,bold=True,color=NAVY)
CF=Font(name=FN,size=10,color="333333"); MF=Font(name=FN,size=10,bold=True,color=NAVY)
TFill=PatternFill("solid",fgColor=NAVY); HFill=PatternFill("solid",fgColor=ACCENT)
MFill=PatternFill("solid",fgColor=ACCENT_LIGHT); BFill=PatternFill("solid",fgColor=LIGHT_GREY)
THIN=Side(style="thin",color="D0D3D8"); BDR=Border(left=THIN,right=THIN,top=THIN,bottom=THIN)
PCT={"Win rate (kumul.)"}
STAGE_PALETTE=[ACCENT,GREEN,GOLD,PLUM,RED,"5B8FB0","8FA998","D0A85C","9A7BAD","C77B7B","4C6B8A","7BA88F"]

# ---------------------------------------------------------------------------
# Pomocné funkce pro grafy (citelnost os, ostre rohy misto oblouck, atd.)
# ---------------------------------------------------------------------------

def _rotate_x_labels(axis, degrees=-45):
    """Natoci popisky osy X, at je jasne, ke kteremu tydnu/datu sloupec patri."""
    bodyPr = RichTextProperties(rot=int(degrees*60000), vert="horz")
    axis.txPr = RichText(
        bodyPr=bodyPr,
        p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=900)), endParaRPr=CharacterProperties(sz=900))],
    )

def _thin_labels(axis, n_weeks, max_labels=16):
    """Kdyz je moc tydnu, nezobrazuj uplne kazdy popisek (jinak se sliji)."""
    if n_weeks > max_labels:
        axis.tickLblSkip = max(1, round(n_weeks / max_labels))

def _sharp_lines(chart):
    """Vypne vyhlazeni (smoothing) carovych grafu - ostre, ne zaoblene prechody."""
    for s in chart.series:
        s.smooth = False

def _style_cat_axis(ch, n_weeks):
    _rotate_x_labels(ch.x_axis)
    _thin_labels(ch.x_axis, n_weeks)
    ch.x_axis.delete = False
    ch.gapWidth = 60

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

def write_raw_block(ws, start_row, n_weeks, stage_order, stage_data):
    """
    Zapise skutecne (nevazene) hodnoty owner x stage x tyden do skrytych
    radku pod viditelnou tabulkou - slouzi jen jako zdroj dat pro grafy
    (funnel, vyvoj funnelu v case), aby grafy mohly ukazovat REALNE castky,
    ne vazene pravdepodobnosti. Radky jsou skryte, uzivatel je v Excelu
    normalne nevidi.
    """
    r = start_row; stage_rows = {}
    for stage in stage_order:
        ws.cell(r,1,f"[raw] {stage}")
        vals = stage_data.get(stage,[0]*n_weeks)
        for w in range(n_weeks):
            ws.cell(r,2+w,vals[w])
        stage_rows[stage] = r; r += 1
    for rr in range(start_row, r):
        ws.row_dimensions[rr].hidden = True
    return stage_rows, r

def _add_charts(ws, anchor_row, header_row, metric_rows, raw_eoy_rows, raw_full_rows, n_weeks, stage_order, week_dates_for_title=None):
    mc = 1+n_weeks
    cats = Reference(ws,min_col=2,max_col=mc,min_row=header_row,max_row=header_row)
    last_week_lbl = week_dates_for_title[-1] if week_dates_for_title else ""

    def line(title, labels, y="Kc"):
        ch=LineChart(); ch.title=title; ch.style=10; ch.y_axis.title=y; ch.height=10; ch.width=26
        for lbl in labels:
            if lbl not in metric_rows: continue
            d=Reference(ws,min_col=1,max_col=mc,min_row=metric_rows[lbl],max_row=metric_rows[lbl])
            ch.add_data(d,titles_from_data=True,from_rows=True)
        ch.set_categories(cats); _sharp_lines(ch); _style_cat_axis(ch,n_weeks); ch.visible_cells_only=False; ch.roundedCorners=False
        return ch

    def bar(title, lbl, color=ACCENT, y="Kc"):
        ch=BarChart(); ch.type="col"; ch.title=title; ch.style=10; ch.y_axis.title=y; ch.height=10; ch.width=26
        d=Reference(ws,min_col=1,max_col=mc,min_row=metric_rows[lbl],max_row=metric_rows[lbl])
        ch.add_data(d,titles_from_data=True,from_rows=True); ch.set_categories(cats)
        ch.series[0].graphicalProperties.solidFill=color; _style_cat_axis(ch,n_weeks); ch.visible_cells_only=False; ch.roundedCorners=False
        return ch

    def funnel_current():
        ch=BarChart(); ch.type="bar"
        sfx=f" ({last_week_lbl})" if last_week_lbl else ""
        ch.title=f"Funnel - rozpad pipeline, skutecne hodnoty, Rolling 18{sfx}"
        ch.style=10; ch.x_axis.title="Kc"; ch.height=10; ch.width=26
        lc=1+n_weeks
        d=Reference(ws,min_col=lc,max_col=lc,min_row=min(raw_full_rows.values()),max_row=max(raw_full_rows.values()))
        cs=Reference(ws,min_col=1,max_col=1,min_row=min(raw_full_rows.values()),max_row=max(raw_full_rows.values()))
        ch.add_data(d,titles_from_data=False); ch.set_categories(cs)
        ch.series[0].graphicalProperties.solidFill=GREEN; ch.legend=None
        ch.dataLabels=DataLabelList(); ch.dataLabels.showVal=True; ch.dataLabels.numFmt="#,##0"
        ch.visible_cells_only=False; ch.roundedCorners=False
        return ch

    def funnel_evolution(title, rows_dict, color_offset=0):
        ch=BarChart(); ch.type="col"; ch.grouping="stacked"; ch.overlap=100
        ch.title=title; ch.style=10; ch.y_axis.title="Kc"; ch.height=10; ch.width=26
        d=Reference(ws,min_col=1,max_col=mc,min_row=min(rows_dict.values()),max_row=max(rows_dict.values()))
        ch.add_data(d,titles_from_data=True,from_rows=True); ch.set_categories(cats)
        for i,s in enumerate(ch.series):
            s.graphicalProperties.solidFill=STAGE_PALETTE[(i+color_offset)%len(STAGE_PALETTE)]
        _style_cat_axis(ch,n_weeks); ch.visible_cells_only=False; ch.roundedCorners=False
        return ch

    def combo_tempo():
        bar_ch=BarChart(); bar_ch.type="col"; bar_ch.style=10
        d=Reference(ws,min_col=1,max_col=mc,min_row=metric_rows["Changes in Rolling 18"],max_row=metric_rows["Changes in Rolling 18"])
        bar_ch.add_data(d,titles_from_data=True,from_rows=True); bar_ch.set_categories(cats)
        bar_ch.series[0].graphicalProperties.solidFill=ACCENT_LIGHT
        bar_ch.y_axis.axId=200; bar_ch.y_axis.title="Zmena Rolling 18 (Kc)"; bar_ch.y_axis.crosses="max"

        line_ch=LineChart(); line_ch.style=10
        for lbl,color in [("Pipeline till end of year",GOLD),("Rolling 18",ACCENT),("Won (kumulativně)",GREEN)]:
            if lbl not in metric_rows: continue
            dd=Reference(ws,min_col=1,max_col=mc,min_row=metric_rows[lbl],max_row=metric_rows[lbl])
            line_ch.add_data(dd,titles_from_data=True,from_rows=True)
        line_ch.set_categories(cats); _sharp_lines(line_ch)
        line_ch.y_axis.axId=100; line_ch.y_axis.title="Kumulativni hodnota (Kc)"

        bar_ch += line_ch
        bar_ch.title="Tempo k cili - pipeline, Rolling 18 a Won v case"
        bar_ch.height=10; bar_ch.width=26
        _style_cat_axis(bar_ch,n_weeks); bar_ch.visible_cells_only=False; bar_ch.roundedCorners=False
        return bar_ch

    charts=[
        bar("Tydenni zmena Rolling 18","Changes in Rolling 18"),
        line("Won vs. Lost (tydne)",["Won","Lost"]),
        funnel_current(),
        combo_tempo(),
        funnel_evolution("Vyvoj funnelu v case - do konce roku (skutecne hodnoty)", raw_eoy_rows),
        funnel_evolution("Vyvoj funnelu v case - Rolling 18 (skutecne hodnoty)", raw_full_rows, color_offset=3),
    ]
    r=anchor_row
    for ch in charts:
        ws.add_chart(ch,f"A{r}")
        r += 22
    return r

METRICS_ORDER=["Won","Lost","Pipeline till end of year","Changes in pipeline till end of the year",
               "Rolling 18","Changes in Rolling 18",
               "Win rate (kumul.)","Prům. velikost dealu","Won (kumulativně)","Goal (kumulativně)"]

def build_person_sheet(wb, owner, sheet_data, stage_order, week_labels_display, week_dates_for_title):
    ws=wb.create_sheet(owner[:31]); n=len(week_labels_display); lc=get_column_letter(1+n)
    goal=sheet_data.get("annual_goal")
    goal_txt=f"Rocni cil: {goal:,.0f} Kc".replace(",","_").replace("_"," ") if goal else "Rocni cil: zatim nestanoven"
    _title(ws,f"{owner} — Sales Pipeline Report",f"Tydenni prehled pipeline, Kc  ·  {goal_txt}",lc)
    hr,sr,mr,er=write_table(ws,4,week_labels_display,stage_order,sheet_data["stage_weighted"],METRICS_ORDER,sheet_data["metrics"])
    raw_eoy_rows, er2 = write_raw_block(ws, er, n, stage_order, sheet_data.get("raw_eoy", {}))
    raw_full_rows, er3 = write_raw_block(ws, er2, n, stage_order, sheet_data.get("raw_full", {}))
    _add_charts(ws,er3+2,hr,mr,raw_eoy_rows,raw_full_rows,n,stage_order,week_dates_for_title)
    ws.freeze_panes="B5"; ws.sheet_view.showGridLines=False

def build_aggregation_sheet(wb, agg_data, stage_order, week_labels_display, week_dates_for_title, leaderboard):
    ws=wb.create_sheet("Aggregation"); n=len(week_labels_display); lc=get_column_letter(1+n)
    _title(ws,"Tapix — Agregovany Sales Pipeline Report","Vsichni obchodnici souctem, Kc",lc)
    hr,sr,mr,er=write_table(ws,4,week_labels_display,stage_order,agg_data["stage_weighted"],METRICS_ORDER,agg_data["metrics"])
    raw_eoy_rows, er2 = write_raw_block(ws, er, n, stage_order, agg_data.get("raw_eoy", {}))
    raw_full_rows, er3 = write_raw_block(ws, er2, n, stage_order, agg_data.get("raw_full", {}))
    end=_add_charts(ws,er3+2,hr,mr,raw_eoy_rows,raw_full_rows,n,stage_order,week_dates_for_title)
    lb=end+2; ws.cell(lb,1,"Zebricek obchodniku (Won celkem)").font=LF; lb+=1
    ws.cell(lb,1,"Obchodnik").font=HF; ws.cell(lb,1).fill=HFill
    ws.cell(lb,2,"Won celkem (Kc)").font=HF; ws.cell(lb,2).fill=HFill
    lbh=lb; lb+=1; lbf=lb
    for name,total in leaderboard:
        ws.cell(lb,1,name).font=CF
        c=ws.cell(lb,2,round(total)); c.number_format="#,##0"; c.font=CF; lb+=1
    lbl=max(lb-1,lbf)
    ch=BarChart(); ch.type="bar"; ch.title="Zebricek obchodniku podle celkoveho Won"; ch.style=10; ch.height=10; ch.width=26
    d=Reference(ws,min_col=2,max_col=2,min_row=lbh,max_row=lbl)
    cs=Reference(ws,min_col=1,max_col=1,min_row=lbf,max_row=lbl)
    ch.add_data(d,titles_from_data=True); ch.set_categories(cs)
    ch.series[0].graphicalProperties.solidFill=ACCENT; ch.legend=None
    ch.dataLabels=DataLabelList(); ch.dataLabels.showVal=True; ch.dataLabels.numFmt="#,##0"
    ws.add_chart(ch,f"D{lbh}"); ws.freeze_panes="B5"; ws.sheet_view.showGridLines=False
    ws.column_dimensions["A"].width=30

def build_ledger_sheet(wb, title, rows, color):
    ws=wb.create_sheet(title)
    _title(ws,f"{title} Deals — automaticky ledger","Generovano tydne z HubSpotu","F")
    headers=["Deal ID","Deal Name","Company","Deal Owner","Close Date","Amount (Kc)","Weeknum"]
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

def build_raw_debug_sheet(wb, title, raw_by_owner, stage_order, week_labels_display, owners):
    ws = wb.create_sheet(title[:31])
    n = len(week_labels_display)
    last_col = get_column_letter(1 + n)
    _title(ws, f"DEBUG: {title}", "Docasne - syrova data BEZ vahy, pro porovnani se starym systemem", last_col)

    r = 4
    for owner in owners:
        c = ws.cell(r, 1, f"— {owner} —")
        c.font = LF
        c.fill = PatternFill("solid", fgColor=ACCENT_LIGHT)
        r += 1
        header_row = r
        ws.cell(r, 1, "Stage").font = HF
        ws.cell(r, 1).fill = HFill
        for w, lbl in enumerate(week_labels_display):
            cc = ws.cell(r, 2 + w, lbl)
            cc.font = HF; cc.fill = HFill; cc.alignment = Alignment(horizontal="center")
        r += 1
        owner_data = raw_by_owner.get(owner, {})
        for stage in stage_order:
            ws.cell(r, 1, stage).font = CF
            vals = owner_data.get(stage, [0] * n)
            for w in range(n):
                cc = ws.cell(r, 2 + w, vals[w])
                cc.number_format = "#,##0"
                cc.font = CF
                cc.alignment = Alignment(horizontal="right")
            r += 1
        r += 1

    for ci in range(1, 2 + n):
        ws.column_dimensions[get_column_letter(ci)].width = 13 if ci > 1 else 32
    ws.freeze_panes = "B5"
    ws.sheet_view.showGridLines = False
    return ws


def new_workbook():
    wb=Workbook(); wb.remove(wb.active); return wb
