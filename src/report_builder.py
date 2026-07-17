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
GOLD="C9A227"; PLUM="6B4C7A"; GRID="E3E6EA"; MUTED="6B7280"; OLIVE="6E8B3D"; WEEKNUM_FILL="EDEFF2"
FN="Calibri"
TF=Font(name=FN,size=16,bold=True,color="FFFFFF"); STF=Font(name=FN,size=10,italic=True,color="FFFFFF")
HF=Font(name=FN,size=10,bold=True,color="FFFFFF"); LF=Font(name=FN,size=10,bold=True,color=NAVY)
CF=Font(name=FN,size=10,color="333333"); MF=Font(name=FN,size=10,bold=True,color=NAVY)
WF=Font(name=FN,size=9,italic=True,color=MUTED)
TFill=PatternFill("solid",fgColor=NAVY); HFill=PatternFill("solid",fgColor=ACCENT)
MFill=PatternFill("solid",fgColor=ACCENT_LIGHT); BFill=PatternFill("solid",fgColor=LIGHT_GREY)
WFill=PatternFill("solid",fgColor=WEEKNUM_FILL)
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

def _axis_label_style(axis, size=900, rotate=None):
    bodyPr = RichTextProperties(rot=int(rotate*60000)) if rotate is not None else RichTextProperties()
    axis.txPr = RichText(
        bodyPr=bodyPr,
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

def _finish(ch, n_weeks, title, y_title="Kč", legend=True, legend_pos="b", rotate_x=-45, thin_x=True):
    """Společný závěrečný 'polish' krok - volá se na konci každého grafu."""
    _set_title(ch, title)
    _axis_label_style(ch.x_axis, rotate=rotate_x)
    if thin_x: _thin_labels(ch.x_axis, n_weeks)
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

def write_table(ws, start_row, week_labels_display, week_nums, stage_order, stage_data, metrics_order, metrics):
    n = len(week_labels_display); r = start_row; header_row = r
    ws.cell(r,1,"Stage / Metrika").font=HF; ws.cell(r,1).fill=HFill; ws.cell(r,1).border=BDR
    for w,lbl in enumerate(week_labels_display):
        c=ws.cell(r,2+w,lbl); c.font=HF; c.fill=HFill; c.alignment=Alignment(horizontal="center"); c.border=BDR
    r += 1
    # Řádek s číslem týdne (ISO Weeknum) - hned pod datem, ať je jasné,
    # o který týden v roce jde (stejné číslování jako v původním excelu).
    weeknum_row = r
    ws.cell(r,1,"Týden č.").font=WF; ws.cell(r,1).fill=WFill; ws.cell(r,1).border=BDR
    for w,wn in enumerate(week_nums):
        c=ws.cell(r,2+w,wn); c.font=WF; c.fill=WFill; c.alignment=Alignment(horizontal="center"); c.border=BDR
    r += 1
    stage_rows={}
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
    return header_row, weeknum_row, stage_rows, metric_rows, r

def write_hidden_block(ws, start_row, n_cols, row_specs):
    """
    Obecná pomocná funkce - zapíše libovolné číselné řady do skrytých
    řádků (slouží jako zdroj dat pro grafy, aniž by to uživatel v Excelu
    viděl). row_specs = list of (label, values[n_cols]).
    Vrací {label: row_index}, next_free_row.
    """
    r = start_row; rows = {}
    for label, values in row_specs:
        ws.cell(r,1,label)
        vals = values if values is not None else [None]*n_cols
        for w in range(n_cols):
            ws.cell(r,2+w, vals[w] if w < len(vals) else None)
        rows[label] = r; r += 1
    for rr in range(start_row, r):
        ws.row_dimensions[rr].hidden = True
    return rows, r

def write_raw_block(ws, start_row, n_weeks, stage_order, stage_data):
    """Skryté řádky se stage x týden hodnotami (raw nebo vážené - podle
    toho, co se do stage_data předá). Používá se jako zdroj dat pro
    funnel grafy."""
    specs = [(f"[data] {stage}", stage_data.get(stage,[0]*n_weeks)) for stage in stage_order]
    rows, next_r = write_hidden_block(ws, start_row, n_weeks, specs)
    return {stage: rows[f"[data] {stage}"] for stage in stage_order}, next_r

def write_chart4_block(ws, start_row, chart4_data):
    """
    Skryté řádky pro graf 4 (Tempo k cíli) - osa X jde přes CELÝ rok
    (weeknum 1..max_week), ne jen zobrazené týdny. Skutečné metriky mají
    hodnotu jen u týdnů, které už nastaly (jinde None -> graf tam prostě
    nekreslí, žádná fabrikovaná data). Goal je lineární přímka po celý rok.
    """
    n = chart4_data["max_week"]
    specs = [
        ("Weeknum", chart4_data["weeknums"]),
        ("Změna Rolling 18 (týdně)", chart4_data["changes_rolling18"]),
        ("Pipeline till end of year", chart4_data["pipeline_eoy"]),
        ("Rolling 18", chart4_data["rolling18"]),
        ("Won (kumulativně)", chart4_data["won_cum"]),
        ("Goal (lineárně, celý rok)", chart4_data["goal_line"]),
    ]
    rows, next_r = write_hidden_block(ws, start_row, n, specs)
    return rows, next_r, n

def _add_charts(ws, anchor_row, header_row, metric_rows, weighted_eoy_rows, raw_full_rows, weighted_full_rows,
                 n_weeks, stage_order, week_dates_for_title=None, chart4_rows=None, chart4_n=None):
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
        # Skládaný sloupcový graf: vývoj funnelu VÁŽENÝCH hodnot v čase,
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
        # "Tempo k cíli" - osa X je CELÝ ROK (weeknum 1..max_week), ne jen
        # zobrazené týdny. Sloupce = týdenní změna Rolling 18 (vedlejší
        # osa), čáry = Pipeline do konce roku, Rolling 18, Won kumulativně
        # (hlavní osa, jen tam, kde už data reálně jsou), a lineární Goal
        # přímka po celý rok (annual_goal / 52 týdnů).
        if not chart4_rows: return None
        n4 = chart4_n; mc4 = 1+n4
        cats4 = Reference(ws,min_col=2,max_col=mc4,min_row=chart4_rows["Weeknum"],max_row=chart4_rows["Weeknum"])

        bar_ch=BarChart(); bar_ch.type="col"
        d=Reference(ws,min_col=1,max_col=mc4,min_row=chart4_rows["Změna Rolling 18 (týdně)"],max_row=chart4_rows["Změna Rolling 18 (týdně)"])
        bar_ch.add_data(d,titles_from_data=True,from_rows=True); bar_ch.set_categories(cats4)
        bar_ch.series[0].graphicalProperties.solidFill=ACCENT_LIGHT
        bar_ch.series[0].graphicalProperties.line.solidFill=ACCENT
        bar_ch.y_axis.axId=200; bar_ch.y_axis.title="Týdenní změna (Kč)"; bar_ch.y_axis.crosses="max"
        bar_ch.y_axis.numFmt=MONEY_FMT; _axis_label_style(bar_ch.y_axis)
        bar_ch.y_axis.majorGridlines=None

        line_ch=LineChart()
        series_specs=[("Pipeline till end of year",GOLD),("Rolling 18",ACCENT),("Won (kumulativně)",GREEN)]
        for name,color in series_specs:
            dd=Reference(ws,min_col=1,max_col=mc4,min_row=chart4_rows[name],max_row=chart4_rows[name])
            line_ch.add_data(dd,titles_from_data=True,from_rows=True)
        for s,(name,color) in zip(line_ch.series, series_specs):
            s.graphicalProperties.line.solidFill = color
            s.marker = Marker(symbol="circle", size=5)
            s.marker.graphicalProperties = GraphicalProperties(solidFill=color, ln=LineProperties(solidFill=color))
        # Goal - samostatná lineární přímka po celý rok, bez značek, čárkovaně,
        # ať je vizuálně jasné, že jde o CÍL, ne o naměřenou hodnotu.
        goal_d=Reference(ws,min_col=1,max_col=mc4,min_row=chart4_rows["Goal (lineárně, celý rok)"],max_row=chart4_rows["Goal (lineárně, celý rok)"])
        line_ch.add_data(goal_d,titles_from_data=True,from_rows=True)
        goal_series = line_ch.series[-1]
        goal_series.graphicalProperties.line.solidFill = OLIVE
        goal_series.graphicalProperties.line.dashStyle = "dash"
        goal_series.marker = Marker(symbol="none")
        line_ch.set_categories(cats4)
        _sharp_lines(line_ch)
        line_ch.y_axis.axId=100; line_ch.y_axis.title="Kumulativní hodnota (Kč)"
        line_ch.y_axis.numFmt=MONEY_FMT; _axis_label_style(line_ch.y_axis)
        if line_ch.y_axis.majorGridlines is not None:
            line_ch.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=GRID, w=6350))

        bar_ch += line_ch
        bar_ch.height=10.5; bar_ch.width=27
        ch = _finish(bar_ch, n4, "Tempo k cíli – pipeline, Rolling 18, Won a Goal po celý rok",
                     y_title=None, legend=True, thin_x=True)
        ch.y_axis.numFmt=MONEY_FMT
        return ch

    charts=[
        bar("Týdenní změna Rolling 18","Changes in Rolling 18", color=ACCENT),
        line("Won vs. Lost (týdně)", [("Won",GREEN),("Lost",RED)]),
        funnel_current(),
        combo_tempo(),
        funnel_evolution("Vývoj funnelu v čase – do konce roku (vážené hodnoty)", weighted_eoy_rows),
        funnel_evolution("Vývoj funnelu v čase – Rolling 18 (vážené hodnoty)", weighted_full_rows, color_offset=3),
    ]
    r=anchor_row
    for ch in charts:
        if ch is None: continue
        ws.add_chart(ch,f"A{r}")
        r += 23
    return r

METRICS_ORDER=["Won","Lost","Pipeline till end of year","Changes in pipeline till end of the year",
               "Rolling 18","Changes in Rolling 18",
               "Win rate (kumul.)","Prům. velikost dealu","Won (kumulativně)","Goal (kumulativně)"]

def build_person_sheet(wb, owner, sheet_data, stage_order, week_labels_display, week_dates_for_title, week_nums):
    ws=wb.create_sheet(owner[:31]); n=len(week_labels_display); lc=get_column_letter(1+n)
    goal=sheet_data.get("annual_goal")
    goal_txt=f"Roční cíl: {goal:,.0f} Kč".replace(",","_").replace("_"," ") if goal else "Roční cíl: zatím nestanoven"
    _title(ws,f"{owner} — Sales Pipeline Report",f"Týdenní přehled pipeline, Kč  ·  {goal_txt}",lc)
    hr,wnr,sr,mr,er=write_table(ws,4,week_labels_display,week_nums,stage_order,sheet_data["stage_weighted"],METRICS_ORDER,sheet_data["metrics"])
    raw_full_rows, er2 = write_raw_block(ws, er, n, stage_order, sheet_data.get("raw_full", {}))
    weighted_full_rows, er3 = write_raw_block(ws, er2, n, stage_order, sheet_data.get("weighted_full", {}))
    chart4_rows, er4, chart4_n = (None, er3, None)
    if sheet_data.get("chart4"):
        chart4_rows, er4, chart4_n = write_chart4_block(ws, er3, sheet_data["chart4"])
    _add_charts(ws,er4+2,hr,mr,sr,raw_full_rows,weighted_full_rows,n,stage_order,week_dates_for_title,chart4_rows,chart4_n)
    ws.freeze_panes="B5"; ws.sheet_view.showGridLines=False

def build_aggregation_sheet(wb, agg_data, stage_order, week_labels_display, week_dates_for_title, leaderboard, week_nums):
    ws=wb.create_sheet("Aggregation"); n=len(week_labels_display); lc=get_column_letter(1+n)
    _title(ws,"Tapix — Agregovaný Sales Pipeline Report","Všichni obchodníci součtem, Kč",lc)
    hr,wnr,sr,mr,er=write_table(ws,4,week_labels_display,week_nums,stage_order,agg_data["stage_weighted"],METRICS_ORDER,agg_data["metrics"])
    raw_full_rows, er2 = write_raw_block(ws, er, n, stage_order, agg_data.get("raw_full", {}))
    weighted_full_rows, er3 = write_raw_block(ws, er2, n, stage_order, agg_data.get("weighted_full", {}))
    chart4_rows, er4, chart4_n = (None, er3, None)
    if agg_data.get("chart4"):
        chart4_rows, er4, chart4_n = write_chart4_block(ws, er3, agg_data["chart4"])
    end=_add_charts(ws,er4+2,hr,mr,sr,raw_full_rows,weighted_full_rows,n,stage_order,week_dates_for_title,chart4_rows,chart4_n)
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
    _title(ws,f"{title} Deals — automatický ledger","Generováno týdně z HubSpotu","G")
    headers=["Deal ID","Deal Name","Company","Deal Owner","Close Date","Amount (Kč)","Týden č. (Weeknum)"]
    r=4
    for ci,h in enumerate(headers):
        c=ws.cell(r,1+ci,h); c.font=HF; c.fill=PatternFill("solid",fgColor=color); c.border=BDR
    r+=1
    for row in sorted(rows, key=lambda x: x[4] if x[4] else date.min):
        for ci,val in enumerate(row):
            c=ws.cell(r,1+ci,val); c.font=CF; c.border=BDR
            if ci==4: c.number_format="dd.mm.yyyy"
            if ci==5: c.number_format="#,##0"
            if ci==6: c.alignment=Alignment(horizontal="center")
        r+=1
    for ci,w in enumerate([10,26,16,14,13,14,16]): ws.column_dimensions[get_column_letter(1+ci)].width=w
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
