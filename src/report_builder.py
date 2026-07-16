"""
Vykreslení finálního .xlsx - design 1:1 podle schváleného mockupu.
Tenhle modul NEPOČÍTÁ žádné metriky (to dělá metrics.py) - jen je vykresluje.
"""
from typing import Dict, List, Optional
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.utils import get_column_letter

NAVY = "1F2A44"
ACCENT = "2E6F9E"
ACCENT_LIGHT = "DCE9F2"
GREEN = "3D8B58"
RED = "B23B3B"
LIGHT_GREY = "F4F5F7"
WHITE = "FFFFFF"

FONT_NAME = "Calibri"
TITLE_FONT = Font(name=FONT_NAME, size=16, bold=True, color=WHITE)
SUBTITLE_FONT = Font(name=FONT_NAME, size=10, italic=True, color=WHITE)
HEADER_FONT = Font(name=FONT_NAME, size=10, bold=True, color=WHITE)
LABEL_FONT = Font(name=FONT_NAME, size=10, bold=True, color=NAVY)
CELL_FONT = Font(name=FONT_NAME, size=10, color="333333")
METRIC_FONT = Font(name=FONT_NAME, size=10, bold=True, color=NAVY)

TITLE_FILL = PatternFill("solid", fgColor=NAVY)
HEADER_FILL = PatternFill("solid", fgColor=ACCENT)
METRIC_FILL = PatternFill("solid", fgColor=ACCENT_LIGHT)
BAND_FILL = PatternFill("solid", fgColor=LIGHT_GREY)

THIN = Side(style="thin", color="D0D3D8")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

PCT_METRICS = {"Win rate (kumul.)"}
NON_MONEY_METRICS = PCT_METRICS


def style_title_block(ws, title, subtitle, last_col_letter):
    ws.merge_cells(f"A1:{last_col_letter}1")
    ws.merge_cells(f"A2:{last_col_letter}2")
    ws["A1"] = title
    ws["A1"].font = TITLE_FONT
    ws["A1"].fill = TITLE_FILL
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws["A2"] = subtitle
    ws["A2"].font = SUBTITLE_FONT
    ws["A2"].fill = TITLE_FILL
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 18


def write_table(
    ws,
    start_row: int,
    week_labels_display: List[str],
    stage_order: List[str],
    stage_data: Dict[str, List[float]],
    metrics_order: List[str],
    metrics: Dict[str, List[float]],
):
    n_weeks = len(week_labels_display)
    r = start_row
    header_row = r
    ws.cell(row=r, column=1, value="Stage / Metrika").font = HEADER_FONT
    ws.cell(row=r, column=1).fill = HEADER_FILL
    ws.cell(row=r, column=1).border = BORDER
    for w, lbl in enumerate(week_labels_display):
        c = ws.cell(row=r, column=2 + w, value=lbl)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center")
        c.border = BORDER
    r += 1

    stage_rows = {}
    for i, stage in enumerate(stage_order):
        band = BAND_FILL if i % 2 == 0 else PatternFill(fill_type=None)
        lbl = ws.cell(row=r, column=1, value=stage)
        lbl.font = CELL_FONT
        lbl.fill = band
        lbl.border = BORDER
        values = stage_data.get(stage, [0] * n_weeks)
        for w in range(n_weeks):
            c = ws.cell(row=r, column=2 + w, value=values[w])
            c.number_format = "#,##0"
            c.font = CELL_FONT
            c.fill = band
            c.border = BORDER
            c.alignment = Alignment(horizontal="right")
        stage_rows[stage] = r
        r += 1

    r += 1
    metric_rows = {}
    for name in metrics_order:
        values = metrics.get(name)
        if values is None:
            continue
        lbl = ws.cell(row=r, column=1, value=name)
        lbl.font = METRIC_FONT
        lbl.fill = METRIC_FILL
        lbl.border = BORDER
        for w in range(n_weeks):
            c = ws.cell(row=r, column=2 + w, value=values[w])
            c.font = METRIC_FONT
            c.fill = METRIC_FILL
            c.border = BORDER
            c.alignment = Alignment(horizontal="right")
            if name in PCT_METRICS:
                c.number_format = "0.0%"
            else:
                c.number_format = "#,##0"
        metric_rows[name] = r
        r += 1

    for c_idx in range(1, 2 + n_weeks):
        ws.column_dimensions[get_column_letter(c_idx)].width = 13 if c_idx > 1 else 30

    return header_row, stage_rows, metric_rows, r


def _end_anchor_row(anchor_row, n_charts):
    rows_used = (n_charts + 1) // 2
    return anchor_row + rows_used * 16


def add_charts(ws, anchor_row, header_row, stage_rows, metric_rows, has_goal, n_weeks, week_dates_for_title=None):
    max_col = 1 + n_weeks
    cats = Reference(ws, min_col=2, max_col=max_col, min_row=header_row, max_row=header_row)

    def make_line(title, rows_labels, y_title="Kč"):
        chart = LineChart()
        chart.title = title
        chart.style = 10
        chart.y_axis.title = y_title
        chart.height = 7.5
        chart.width = 15
        for label in rows_labels:
            if label not in metric_rows:
                continue
            row = metric_rows[label]
            data = Reference(ws, min_col=1, max_col=max_col, min_row=row, max_row=row)
            chart.add_data(data, titles_from_data=True, from_rows=True)
        chart.set_categories(cats)
        return chart

    def make_bar(title, row_label, color=ACCENT, y_title="Kč"):
        chart = BarChart()
        chart.type = "col"
        chart.title = title
        chart.style = 10
        chart.y_axis.title = y_title
        chart.height = 7.5
        chart.width = 15
        row = metric_rows[row_label]
        data = Reference(ws, min_col=1, max_col=max_col, min_row=row, max_row=row)
        chart.add_data(data, titles_from_data=True, from_rows=True)
        chart.set_categories(cats)
        chart.series[0].graphicalProperties.solidFill = color
        return chart

    def make_funnel_latest_week():
        chart = BarChart()
        chart.type = "bar"
        title_suffix = f" ({week_dates_for_title[-1]})" if week_dates_for_title else ""
        chart.title = f"Funnel – rozpad pipeline podle stage{title_suffix}"
        chart.style = 10
        chart.x_axis.title = "Kč"
        chart.height = 7.8
        chart.width = 15
        last_col = 1 + n_weeks
        data = Reference(ws, min_col=last_col, max_col=last_col,
                          min_row=min(stage_rows.values()), max_row=max(stage_rows.values()))
        cats_stage = Reference(ws, min_col=1, max_col=1,
                                min_row=min(stage_rows.values()), max_row=max(stage_rows.values()))
        chart.add_data(data, titles_from_data=False)
        chart.set_categories(cats_stage)
        chart.series[0].graphicalProperties.solidFill = GREEN
        chart.legend = None
        return chart

    charts = [
        make_bar("Týdenní změna pipeline", "Changes in pipeline"),
        make_line("Won vs. Lost (týdně)", ["Won", "Lost"]),
        make_funnel_latest_week(),
    ]
    if has_goal:
        charts.append(make_line("Tempo k cíli – kumulativní Won vs. Goal", ["Won (kumulativně)", "Goal (kumulativně)"]))
    charts.append(make_line("Win rate v čase", ["Win rate (kumul.)"], y_title="%"))
    charts.append(make_line("Průměrná velikost dealu", ["Prům. velikost dealu"]))

    r = anchor_row
    col_positions = ["A", "H"]
    for i, chart in enumerate(charts):
        anchor = f"{col_positions[i % 2]}{r}"
        ws.add_chart(chart, anchor)
        if i % 2 == 1:
            r += 16
    return _end_anchor_row(anchor_row, len(charts))


def build_person_sheet(wb, owner: str, sheet_data: dict, stage_order: List[str], week_labels_display: List[str],
                        week_dates_for_title: List[str]):
    ws = wb.create_sheet(owner[:31])
    n_weeks = len(week_labels_display)
    last_col_letter = get_column_letter(1 + n_weeks)

    goal = sheet_data.get("annual_goal")
    goal_txt = f"Roční cíl: {goal:,.0f} Kč".replace(",", " ") if goal else "Roční cíl: zatím nestanoven"
    style_title_block(ws, f"{owner} — Sales Pipeline Report",
                       f"Týdenní přehled pipeline, Kč  ·  {goal_txt}", last_col_letter)

    metrics_order = [
        "Won", "Lost", "Pipeline till end of year", "Changes in pipeline", "Rolling 18",
        "Win rate (kumul.)", "Prům. velikost dealu", "Won (kumulativně)", "Goal (kumulativně)",
    ]
    header_row, stage_rows, metric_rows, end_row = write_table(
        ws, start_row=4, week_labels_display=week_labels_display, stage_order=stage_order,
        stage_data=sheet_data["stage_weighted"], metrics_order=metrics_order, metrics=sheet_data["metrics"],
    )
    has_goal = sheet_data["metrics"].get("Goal (kumulativně)") is not None
    add_charts(ws, anchor_row=end_row + 2, header_row=header_row, stage_rows=stage_rows,
               metric_rows=metric_rows, has_goal=has_goal, n_weeks=n_weeks,
               week_dates_for_title=week_dates_for_title)
    ws.freeze_panes = "B5"
    ws.sheet_view.showGridLines = False
    return ws


def build_aggregation_sheet(wb, agg_data: dict, stage_order: List[str], week_labels_display: List[str],
                             week_dates_for_title: List[str], leaderboard: List[tuple]):
    ws2 = wb.create_sheet("Aggregation")
    n_weeks = len(week_labels_display)
    last_col_letter = get_column_letter(1 + n_weeks)

    style_title_block(ws2, "Tapix — Agregovaný Sales Pipeline Report",
                       "Všichni obchodníci součtem, Kč", last_col_letter)

    metrics_order = [
        "Won", "Lost", "Pipeline till end of year", "Changes in pipeline", "Rolling 18",
        "Win rate (kumul.)", "Prům. velikost dealu", "Won (kumulativně)", "Goal (kumulativně)",
    ]
    header_row, stage_rows, metric_rows, end_row = write_table(
        ws2, start_row=4, week_labels_display=week_labels_display, stage_order=stage_order,
        stage_data=agg_data["stage_weighted"], metrics_order=metrics_order, metrics=agg_data["metrics"],
    )
    has_goal = agg_data["metrics"].get("Goal (kumulativně)") is not None
    end_after_charts = add_charts(ws2, anchor_row=end_row + 2, header_row=header_row, stage_rows=stage_rows,
                                   metric_rows=metric_rows, has_goal=has_goal, n_weeks=n_weeks,
                                   week_dates_for_title=week_dates_for_title)

    lb_row = end_after_charts + 2
    ws2.cell(row=lb_row, column=1, value="Žebříček obchodníků (Won celkem)").font = LABEL_FONT
    lb_row += 1
    ws2.cell(row=lb_row, column=1, value="Obchodník").font = HEADER_FONT
    ws2.cell(row=lb_row, column=1).fill = HEADER_FILL
    ws2.cell(row=lb_row, column=2, value="Won celkem (Kč)").font = HEADER_FONT
    ws2.cell(row=lb_row, column=2).fill = HEADER_FILL
    lb_header_row = lb_row
    lb_row += 1
    lb_first_data_row = lb_row
    for name, total in leaderboard:
        ws2.cell(row=lb_row, column=1, value=name).font = CELL_FONT
        c = ws2.cell(row=lb_row, column=2, value=round(total))
        c.number_format = "#,##0"
        c.font = CELL_FONT
        lb_row += 1
    lb_last_data_row = max(lb_row - 1, lb_first_data_row)

    chart_lb = BarChart()
    chart_lb.type = "bar"
    chart_lb.title = "Žebříček obchodníků podle celkového Won"
    chart_lb.style = 10
    chart_lb.height = 8
    chart_lb.width = 15
    data_lb = Reference(ws2, min_col=2, max_col=2, min_row=lb_header_row, max_row=lb_last_data_row)
    cats_lb = Reference(ws2, min_col=1, max_col=1, min_row=lb_first_data_row, max_row=lb_last_data_row)
    chart_lb.add_data(data_lb, titles_from_data=True)
    chart_lb.set_categories(cats_lb)
    chart_lb.series[0].graphicalProperties.solidFill = ACCENT
    chart_lb.legend = None
    ws2.add_chart(chart_lb, f"D{lb_header_row}")

    ws2.freeze_panes = "B5"
    ws2.sheet_view.showGridLines = False
    ws2.column_dimensions["A"].width = 30


def build_ledger_sheet(wb, title: str, rows: List[list], color: str):
    """rows: [deal_id, deal_name, company_name, owner_name, closedate(date), amount, weeknum]"""
    ws = wb.create_sheet(title)
    style_title_block(ws, f"{title} Deals — automatický ledger",
                       "Generováno týdně z HubSpotu (nahrazuje ruční doplňování)", "F")
    headers = ["Deal ID", "Deal Name", "Company", "Deal Owner", "Close Date", "Amount (Kč)", "Weeknum"]
    r = 4
    for c_i, h in enumerate(headers):
        c = ws.cell(row=r, column=1 + c_i, value=h)
        c.font = HEADER_FONT
        c.fill = PatternFill("solid", fgColor=color)
        c.border = BORDER
    r += 1
    rows_sorted = sorted(rows, key=lambda x: x[4] if x[4] else date.min)
    for row in rows_sorted:
        for c_i, val in enumerate(row):
            c = ws.cell(row=r, column=1 + c_i, value=val)
            c.font = CELL_FONT
            c.border = BORDER
            if c_i == 4:
                c.number_format = "dd.mm.yyyy"
            if c_i == 5:
                c.number_format = "#,##0"
        r += 1
    widths = [10, 26, 16, 14, 13, 14, 9]
    for c_i, wdt in enumerate(widths):
        ws.column_dimensions[get_column_letter(1 + c_i)].width = wdt
    ws.freeze_panes = "A5"
    ws.sheet_view.showGridLines = False
    return ws


def new_workbook() -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)
    return wb
