from typing import Dict, List, Optional
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.text import RichText
from openpyxl.chart.legend import Legend
from openpyxl.chart.marker import Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import RichTextProperties, Paragraph, ParagraphProperties, CharacterProperties
from openpyxl.utils import get_column_letter

NAVY="1F2A44"; ACCENT="2E6F9E"; ACCENT_LIGHT="DCE9F2"; GREEN="3D8B58"; RED="B23B3B"; LIGHT_GREY="F4F5F7"
GOLD="C9A227"; PLUM="6B4C7A"; GRID="E3E6EA"; MUTED="6B7280"
FN="Calibri"
TF=Font(name=FN,size=16,bold=True,color="FFFFFF"); STF=Font(name=FN,size=10,italic=True,color="FFFFFF")
HF=Font(name=FN,size=10,bold=True,color="FFFFFF"); LF=Font(name=FN,size=10,bold=True,color=NAVY)
CF=Font(name=FN,size=10,color="333333"); MF=Font(name=FN,size=10,bold=True,color=NAVY)
TFill=PatternFill("solid",fgColor=NAVY); HFill=PatternFill("solid",fgColor=ACCENT)
MFill=PatternFill("solid",fgColor=ACCENT_LIGHT); BFill=PatternFill("solid",fgColor=LIGHT_GREY)
THIN=Side(style="thin",color="D0D3D8"); BDR=Border(left=THIN,right=THIN,top=THIN,bottom=THIN)
PCT={"Win rate (kumul.)"}
STAGE_PALETTE=[ACCENT,GREEN,GOLD,PLUM,RED,"5B8FB0","8FA998","D0A85C","9A7BAD","C77B7B","4C6B8A","7BA88F"]
MONEY_FMT='#,##0" Kč"'

# ---------------------------------------------------------------------------
# Pomocné funkce pro vzhled grafů - čitelnost os, ostré rohy, jednotné barvy,
# jednotný styl titulků/legendy, bez zbytečného rámování.
# ---------------------------------------------------------------------------

def _rotate_x_labels(axis, degrees=-45):
    """Natočí popisky osy X, ať je jasné, ke kterému týdnu/datu sloupec patří."""
    bodyPr = RichTextProperties(rot=int(degrees*60000), vert="horz")
    axis.txPr = RichText(
        bodyPr=bodyPr,
        p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=900,solidFill=MUTED)),
                      endParaRPr=CharacterProperties(sz=900,solidFill=MUTED))],
    )

def _axis_label_style(axis, size=900):
    axis.txPr = RichText(
        bodyPr=RichTextProperties(),
        p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=size,solidFill=MUTED)),
                      endParaRPr=CharacterProperties(sz=size,solidFill=MUTED))],
    )

def _thin_labels(axis, n_weeks, max_labels=16):
    """Když je moc týdnů, nezobrazuj úplně každý popisek (jinak se slijou)."""
    if n_weeks > max_labels:
        axis.tickLblSkip = max(1, round(n_weeks / max_labels))

def _sharp_lines(chart, width_pt=2.25):
    """Vypne vyhlazení (smoothing) čárových grafů - ostré, ne zaoblené přechody,
    a ztlustí čáru, ať je na první pohled dobře vidět."""
    for s in chart.series:
        s.smooth = False
        s.graphicalProperties.line.width = int(width_pt*12700)

def _set_title(ch, text, size=1400, color=NAVY):
    ch.title = text
    try:
        run = ch.title.tx.rich.p[0].r[0]
        run.rPr = CharacterProperties(sz=size, b=True, solidFill=color, latin=None)
    except Exception:
        pass

def _no_border(ch):
    """Odstraní rámeček kolem celého grafu - čistší, 'plovoucí' vzhled."""
    ch.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))

def _style_value_axis(ch, title=None, fmt=MONEY_FMT):
    if title: ch.y_axis.title = title
    ch.y_axis.numFmt = fmt
    if ch.y_axis.majorGridlines is not None:
        ch.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=GRID, w=6350))
    _axis_label_style(ch.y_axis)
    ch.y_axis.delete = False

def _style_legend(ch, position="b", size=850):
    ch.legend = Legend(); ch.legend.position = position; ch.legend.overlay = False
    ch.legend.txPr = RichText(
        bodyPr=RichTextProperties(),
        p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=size,solidFill=MUTED)), endParaRPr=None)],
    )

def _finish(ch, n_weeks, title, y_title="Kč", legend=True, legend_pos="b"):
    """Společný závěrečný 'polish' krok - volá se na konci každého grafu."""
    _set_title(ch, title)
    _rotate_x_labels(ch.x_axis)
    _thin_labels(ch.x_axis, n_weeks)
    _axis_label_style(ch.x_axis)
    ch.x_axis.delete = False
    ch.gapWidth = 55
    _style_value_axis(ch, y_title)
    if legend: _style_legend(ch, legend_pos)
    else: ch.legend = None
    _no_border(ch)
    ch.roundedCorners = False
    ch.visible_cells_only = False
    return ch

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
    Zapíše skutečné (nevážené) hodnoty owner x stage x týden do skrytých
    řádků pod viditelnou tabulkou - slouží jen jako zdroj dat pro grafy
    (funnel, vývoj funnelu v čase), aby grafy mohly ukazovat REÁLNÉ částky,
    ne vážené pravděpodobností. Řádky jsou skryté, uživatel je v Excelu
    normálně nevidí.
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

    def line(title, series_colors, y="Kč"):
        # series_colors: list of (metric_label, hex_color)
        ch=LineChart(); ch.height=10.5; ch.width=27
        for lbl,color in series_colors:
            if lbl not in metric_rows: continue
            d=Reference(ws,min_col=1,max_col=mc,min_row=metric_rows[lbl],max_row=metric_rows[lbl])
            ch.add_data(d,titles_from_data=True,from_rows=True)
        ch.set_categories(cats)
        for s,(lbl,color) in zip(ch.series, [x for x in series_colors if x[0] in metric_rows]):
            s.graphicalProperties.line.solidFill = color
            s.marker = Marker(symbol="circle", size=5)
            s.marker.graphicalProperties = GraphicalProperties(solidFill=color, ln=LineProperties(solidFill=color))
        _sharp_lines(ch)
        return _finish(ch, n_weeks, title, y)

    def bar(title, lbl, color=ACCENT, y="Kč"):
        ch=BarChart(); ch.type="col"; ch.height=10.5; ch.width=27
        d=Reference(ws,min_col=1,max_col=mc,min_row=metric_rows[lbl],max_row=metric_rows[lbl])
        ch.add_data(d,titles_from_data=True,from_rows=True); ch.set_categories(cats)
        ch.series[0].graphicalProperties.solidFill=color
        ch.series[0].graphicalProperties.line.noFill=True
        return _finish(ch, n_weeks, title, y, legend=False)

    def funnel_current():
        # Aktuální rozpad pipeline (poslední týden) - SKUTEČNÉ (nevážené)
        # hodnoty, na základě Rolling 18 datasetu (ne "do konce roku").
        ch=BarChart(); ch.type="bar"; ch.height=10.5; ch.width=27
        lc=1+n_weeks
        min_r, max_r = min(raw_full_rows.values()), max(raw_full_rows.values())
        d=Reference(ws,min_col=lc,max_col=lc,min_row=min_r,max_row=max_r)
        cs=Reference(ws,min_col=1,max_col=1,min_row=min_r,max_row=max_r)
        ch.add_data(d,titles_from_data=False); ch.set_categories(cs)
        ch.series[0].graphicalProperties.solidFill=GREEN
        ch.series[0].graphicalProperties.line.noFill=True
        ch.dataLabels=DataLabelList(); ch.dataLabels.showVal=True; ch.dataLabels.numFmt='#,##0" Kč"'
        ch.dataLabels.txPr = RichText(bodyPr=RichTextProperties(),
            p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=850,b=True,solidFill=NAVY)), endParaRPr=None)])
        # Fáze na začátku funnelu (Lead Engaged...) nahoru, ne na spodek.
        ch.y_axis.scaling.orientation = "maxMin"
        sfx=f" ({last_week_lbl})" if last_week_lbl else ""
        ch = _finish(ch, n_weeks, f"Funnel – rozpad pipeline, Rolling 18{sfx}", "Kč", legend=False)
        ch.x_axis.title = None
        return ch

    def funnel_evolution(title, rows_dict, color_offset=0):
        # Skládaný sloupcový graf: vývoj funnelu (skutečné hodnoty) v čase,
        # po týdnech, od začátku roku - jedna série na fázi.
        ch=BarChart(); ch.type="col"; ch.grouping="stacked"; ch.overlap=100; ch.height=11; ch.width=27
        d=Reference(ws,min_col=1,max_col=mc,min_row=min(rows_dict.values()),max_row=max(rows_dict.values()))
        ch.add_data(d,titles_from_data=True,from_rows=True); ch.set_categories(cats)
        for i,s in enumerate(ch.series):
            color = STAGE_PALETTE[(i+color_offset)%len(STAGE_PALETTE)]
            s.graphicalProperties.solidFill = color
            s.graphicalProperties.line.noFill = True
        return _finish(ch, n_weeks, title, "Kč", legend=True)

    def combo_tempo():
        # Náhrada za "Tempo k cíli" - kombinovaný graf: sloupce = týdenní
        # změna Rolling 18 (vedlejší osa), čáry = Pipeline do konce roku,
        # Rolling 18, Won kumulativně (hlavní osa).
        bar_ch=BarChart(); bar_ch.type="col"
        d=Reference(ws,min_col=1,max_col=mc,min_row=metric_rows["Changes in Rolling 18"],max_row=metric_rows["Changes in Rolling 18"])
        bar_ch.add_data(d,titles_from_data=True,from_rows=True); bar_ch.set_categories(cats)
        bar_ch.series[0].graphicalProperties.solidFill=ACCENT_LIGHT
        bar_ch.series[0].graphicalProperties.line.solidFill=ACCENT
        bar_ch.y_axis.axId=200; bar_ch.y_axis.title="Týdenní změna (Kč)"; bar_ch.y_axis.crosses="max"
        bar_ch.y_axis.numFmt=MONEY_FMT; _axis_label_style(bar_ch.y_axis)
        bar_ch.y_axis.majorGridlines=None

        line_ch=LineChart()
        series_colors=[("Pipeline till end of year",GOLD),("Rolling 18",ACCENT),("Won (kumulativně)",GREEN)]
        for lbl,color in series_colors:
            if lbl not in metric_rows: continue
            dd=Reference(ws,min_col=1,max_col=mc,min_row=metric_rows[lbl],max_row=metric_rows[lbl])
            line_ch.add_data(dd,titles_from_data=True,from_rows=True)
        line_ch.set_categories(cats)
        for s,(lbl,color) in zip(line_ch.series, [x for x in series_colors if x[0] in metric_rows]):
            s.graphicalProperties.line.solidFill = color
            s.marker = Marker(symbol="circle", size=5)
            s.marker.graphicalProperties = GraphicalProperties(solidFill=color, ln=LineProperties(solidFill=color))
        _sharp_lines(line_ch)
        line_ch.y_axis.axId=100; line_ch.y_axis.title="Kumulativní hodnota (Kč)"
        line_ch.y_axis.numFmt=MONEY_FMT; _axis_label_style(line_ch.y_axis)
        if line_ch.y_axis.majorGridlines is not None:
            line_ch.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=GRID, w=6350))

        bar_ch += line_ch
        bar_ch.height=10.5; bar_ch.width=27
        ch = _finish(bar_ch, n_weeks, "Tempo k cíli – pipeline, Rolling 18 a Won v čase", y_title=None, legend=True)
        ch.y_axis.numFmt=MONEY_FMT
        return ch

    charts=[
        bar("Týdenní změna Rolling 18","Changes in Rolling 18", color=ACCENT),
        line("Won vs. Lost (týdně)", [("Won",GREEN),("Lost",RED)]),
        funnel_current(),
        combo_tempo(),
        funnel_evolution("Vývoj funnelu v čase – do konce roku (skutečné hodnoty)", raw_eoy_rows),
        funnel_evolution("Vývoj funnelu v čase – Rolling 18 (skutečné hodnoty)", raw_full_rows, color_offset=3),
    ]
    r=anchor_row
    for ch in charts:
        ws.add_chart(ch,f"A{r}")
        r += 23
    return r

METRICS_ORDER=["Won","Lost","Pipeline till end of year","Changes in pipeline till end of the year",
               "Rolling 18","Changes in Rolling 18",
               "Win rate (kumul.)","Prům. velikost dealu","Won (kumulativně)","Goal (kumulativně)"]

def build_person_sheet(wb, owner, sheet_data, stage_order, week_labels_display, week_dates_for_title):
    ws=wb.create_sheet(owner[:31]); n=len(week_labels_display); lc=get_column_letter(1+n)
    goal=sheet_data.get("annual_goal")
    goal_txt=f"Roční cíl: {goal:,.0f} Kč".replace(",","_").replace("_"," ") if goal else "Roční cíl: zatím nestanoven"
    _title(ws,f"{owner} — Sales Pipeline Report",f"Týdenní přehled pipeline, Kč  ·  {goal_txt}",lc)
    hr,sr,mr,er=write_table(ws,4,week_labels_display,stage_order,sheet_data["stage_weighted"],METRICS_ORDER,sheet_data["metrics"])
    raw_eoy_rows, er2 = write_raw_block(ws, er, n, stage_order, sheet_data.get("raw_eoy", {}))
    raw_full_rows, er3 = write_raw_block(ws, er2, n, stage_order, sheet_data.get("raw_full", {}))
    _add_charts(ws,er3+2,hr,mr,raw_eoy_rows,raw_full_rows,n,stage_order,week_dates_for_title)
    ws.freeze_panes="B5"; ws.sheet_view.showGridLines=False

def build_aggregation_sheet(wb, agg_data, stage_order, week_labels_display, week_dates_for_title, leaderboard):
    ws=wb.create_sheet("Aggregation"); n=len(week_labels_display); lc=get_column_letter(1+n)
    _title(ws,"Tapix — Agregovaný Sales Pipeline Report","Všichni obchodníci součtem, Kč",lc)
    hr,sr,mr,er=write_table(ws,4,week_labels_display,stage_order,agg_data["stage_weighted"],METRICS_ORDER,agg_data["metrics"])
    raw_eoy_rows, er2 = write_raw_block(ws, er, n, stage_order, agg_data.get("raw_eoy", {}))
    raw_full_rows, er3 = write_raw_block(ws, er2, n, stage_order, agg_data.get("raw_full", {}))
    end=_add_charts(ws,er3+2,hr,mr,raw_eoy_rows,raw_full_rows,n,stage_order,week_dates_for_title)
    lb=end+2; ws.cell(lb,1,"Žebříček obchodníků (Won celkem)").font=LF; lb+=1
    ws.cell(lb,1,"Obchodník").font=HF; ws.cell(lb,1).fill=HFill
    ws.cell(lb,2,"Won celkem (Kč)").font=HF; ws.cell(lb,2).fill=HFill
    lbh=lb; lb+=1; lbf=lb
    for name,total in leaderboard:
        ws.cell(lb,1,name).font=CF
        c=ws.cell(lb,2,round(total)); c.number_format="#,##0"; c.font=CF; lb+=1
    lbl=max(lb-1,lbf)
    ch=BarChart(); ch.type="bar"; ch.height=10.5; ch.width=27
    d=Reference(ws,min_col=2,max_col=2,min_row=lbh,max_row=lbl)
    cs=Reference(ws,min_col=1,max_col=1,min_row=lbf,max_row=lbl)
    ch.add_data(d,titles_from_data=True); ch.set_categories(cs)
    ch.series[0].graphicalProperties.solidFill=ACCENT
    ch.series[0].graphicalProperties.line.noFill=True
    ch.dataLabels=DataLabelList(); ch.dataLabels.showVal=True; ch.dataLabels.numFmt=MONEY_FMT
    ch.y_axis.scaling.orientation="maxMin"
    ch = _finish(ch, len(leaderboard) or 1, "Žebříček obchodníků podle celkového Won", "Kč", legend=False)
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

def build_raw_debug_sheet(wb, title, raw_by_owner, stage_order, week_labels_display, owners):
    ws = wb.create_sheet(title[:31])
    n = len(week_labels_display)
    last_col = get_column_letter(1 + n)
    _title(ws, f"DEBUG: {title}", "Dočasné - syrová data BEZ váhy, pro porovnání se starým systémem", last_col)

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
