"""Сборка Excel-отчёта: данные, сводка, разрезы по месяцам/менеджерам/категориям/клиентам, проблемы."""
from __future__ import annotations

import datetime as dt
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .clean import CleanResult

HEAD_FILL = PatternFill("solid", fgColor="2F5D50")
HEAD_FONT = Font(bold=True, color="FFFFFF")
THIN = Side(style="thin", color="D9D9D9")
MONEY = '#,##0.00'
DATA = "Данные"


def _header(ws, row: int, titles: list[str]) -> None:
    for c, t in enumerate(titles, start=1):
        cell = ws.cell(row=row, column=c, value=t)
        cell.fill, cell.font = HEAD_FILL, HEAD_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=THIN)


def _widths(ws, widths: list[int]) -> None:
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def month_key(d: dt.date) -> str:
    return f"{d.year:04d}-{d.month:02d}"


def months_between(a: dt.date, b: dt.date) -> list[str]:
    out, y, m = [], a.year, a.month
    while (y, m) <= (b.year, b.month):
        out.append(f"{y:04d}-{m:02d}")
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return out


def build_report(result: CleanResult, path, title: str = "Отчёт по заказам") -> None:
    orders = sorted(result.orders, key=lambda o: (o.date, o.order))
    last = max(len(orders) + 1, 2)
    rng = lambda col: f"{DATA}!${col}$2:${col}${last}"
    wb = Workbook()
    summary = wb.active
    summary.title = "Сводка"
    by_month = wb.create_sheet("По месяцам")
    by_mgr = wb.create_sheet("По менеджерам")
    by_cat = wb.create_sheet("По категориям")
    top = wb.create_sheet("Топ клиентов")
    data = wb.create_sheet(DATA)
    issues = wb.create_sheet("Проблемы")

    # --- Данные: сумма считается формулой, чтобы отчёт оставался «живым» ---
    _header(data, 1, ["Номер заказа", "Дата", "Месяц", "Клиент", "Менеджер", "Категория",
                      "Количество", "Цена", "Сумма"])
    for i, o in enumerate(orders, start=2):
        data.append([o.order, o.date, month_key(o.date), o.client, o.manager, o.category,
                     o.qty, o.price, f"=G{i}*H{i}"])
        data.cell(row=i, column=2).number_format = "DD.MM.YYYY"
        data.cell(row=i, column=8).number_format = MONEY
        data.cell(row=i, column=9).number_format = MONEY
    data.freeze_panes = "A2"
    data.auto_filter.ref = f"A1:I{last}"
    _widths(data, [14, 12, 10, 24, 20, 22, 12, 14, 16])

    # --- Разрезы: формулы SUMIFS/COUNTIFS по листу «Данные» ---
    total_rev = f"SUM({rng('I')})"

    def breakdown(ws, name: str, col: str, keys: list[str], chart_title: str) -> None:
        _header(ws, 1, [name, "Выручка", "Заказов", "Доля"])
        for i, k in enumerate(keys, start=2):
            ws.cell(row=i, column=1, value=k)
            ws.cell(row=i, column=2, value=f"=SUMIFS({rng('I')},{rng(col)},A{i})").number_format = MONEY
            ws.cell(row=i, column=3, value=f"=COUNTIFS({rng(col)},A{i})")
            ws.cell(row=i, column=4, value=f"=IF({total_rev}=0,0,B{i}/{total_rev})").number_format = "0.0%"
        _widths(ws, [26, 16, 10, 10])
        ws.freeze_panes = "A2"
        if keys:
            ch = BarChart()
            ch.title, ch.legend = chart_title, None
            ch.add_data(Reference(ws, min_col=2, min_row=1, max_row=len(keys) + 1), titles_from_data=True)
            ch.set_categories(Reference(ws, min_col=1, min_row=2, max_row=len(keys) + 1))
            ch.height, ch.width = 8, 18
            ws.add_chart(ch, "F2")

    months = []
    if orders:
        months = months_between(orders[0].date, orders[-1].date)
    breakdown(by_month, "Месяц", "C", months, "Выручка по месяцам")
    e1 = by_month.cell(row=1, column=5, value="К пред. месяцу")
    e1.fill, e1.font = HEAD_FILL, HEAD_FONT
    e1.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i in range(3, len(months) + 2):
        by_month.cell(row=i, column=5, value=f'=IF(B{i-1}=0,"",B{i}/B{i-1}-1)').number_format = "+0.0%;-0.0%"
    for col, ws, name, title2 in (("E", by_mgr, "Менеджер", "Выручка по менеджерам"),
                                  ("F", by_cat, "Категория", "Выручка по категориям")):
        rev = defaultdict(float)
        for o in orders:
            rev[o.manager if col == "E" else o.category] += o.amount
        keys = sorted(rev, key=lambda k: (-rev[k], k))
        breakdown(ws, name, col, keys, title2)
    if months:  # график по месяцам сдвинут правее колонки динамики
        by_month._charts[0].anchor = "G2"

    # --- Топ клиентов ---
    crev = defaultdict(float)
    for o in orders:
        crev[o.client] += o.amount
    clients = sorted(crev, key=lambda k: (-crev[k], k))[:10]
    _header(top, 1, ["№", "Клиент", "Выручка", "Заказов", "Доля"])
    for i, k in enumerate(clients, start=2):
        top.cell(row=i, column=1, value=i - 1)
        top.cell(row=i, column=2, value=k)
        top.cell(row=i, column=3, value=f"=SUMIFS({rng('I')},{rng('D')},B{i})").number_format = MONEY
        top.cell(row=i, column=4, value=f"=COUNTIFS({rng('D')},B{i})")
        top.cell(row=i, column=5, value=f"=IF({total_rev}=0,0,C{i}/{total_rev})").number_format = "0.0%"
    _widths(top, [5, 28, 16, 10, 10])

    # --- Проблемы ---
    _header(issues, 1, ["Строка в файле", "Что сделано", "Причина", "Исходная строка"])
    for it in result.issues:
        issues.append([it.row, it.kind, it.reason, it.raw])
    _widths(issues, [14, 14, 60, 70])
    issues.freeze_panes = "A2"

    # --- Сводка ---
    summary["A1"] = title
    summary["A1"].font = Font(bold=True, size=16)
    rows = [
        ("Выручка", f"={total_rev}", MONEY),
        ("Заказов", f"=COUNTA({rng('A')})", "0"),
        ("Средний чек", "=IF(B4=0,0,B3/B4)", MONEY),
        ("Первый заказ", f"=IF(B4=0,\"\",MIN({rng('B')}))", "DD.MM.YYYY"),
        ("Последний заказ", f"=IF(B4=0,\"\",MAX({rng('B')}))", "DD.MM.YYYY"),
        ("Лучший месяц", f"=IF(B4=0,\"\",INDEX('По месяцам'!$A$2:$A${max(len(months),1)+1},"
                         f"MATCH(MAX('По месяцам'!$B$2:$B${max(len(months),1)+1}),"
                         f"'По месяцам'!$B$2:$B${max(len(months),1)+1},0)))", "General"),
        ("Прочитано строк", result.total_rows, "0"),
        ("Удалено дубликатов", result.duplicates, "0"),
        ("Отброшено строк с ошибками", result.rejected, "0"),
    ]
    for r, (label, value, fmt) in enumerate(rows, start=3):
        summary.cell(row=r, column=1, value=label).font = Font(bold=True)
        c = summary.cell(row=r, column=2, value=value)
        c.number_format = fmt
        c.alignment = Alignment(horizontal="right")
    summary["A13"] = "Подробности об отброшенных строках — на листе «Проблемы»."
    summary["A13"].font = Font(italic=True, color="6B6A64")
    _widths(summary, [40, 22])
    by_month.column_dimensions["E"].width = 18
    for ws in wb.worksheets:                      # при печати лист целиком по ширине страницы
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)
