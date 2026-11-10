from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .clean import clean_rows, read_table
from .demo import generate
from .report import build_report


def main(argv=None) -> int:
    p = argparse.ArgumentParser(prog="xlreport", description="Выгрузка заказов (CSV/XLSX) → Excel-отчёт")
    sub = p.add_subparsers(dest="cmd", required=True)
    b = sub.add_parser("build", help="собрать отчёт из выгрузки")
    b.add_argument("input", help="файл .csv или .xlsx")
    b.add_argument("-o", "--out", default="report.xlsx")
    b.add_argument("--title", default="Отчёт по заказам")
    d = sub.add_parser("demo", help="создать пример «грязной» выгрузки")
    d.add_argument("-o", "--out", default="orders.csv")
    d.add_argument("-n", type=int, default=500, help="сколько заказов")
    d.add_argument("--seed", type=int, default=0)
    a = p.parse_args(argv)

    if a.cmd == "demo":
        demo = generate(a.n, a.seed)
        Path(a.out).write_text(demo.csv_text, encoding="utf-8-sig")
        print(f"{a.out}: {a.n} заказов + {demo.n_duplicates} дубликатов, "
              f"{demo.n_conflicts} конфликтов номера, {demo.n_bad} битых строк")
        return 0
    try:
        res = clean_rows(read_table(a.input))
    except (ValueError, OSError) as e:
        print(f"Ошибка: {e}", file=sys.stderr)
        return 2
    build_report(res, a.out, a.title)
    print(f"Прочитано строк: {res.total_rows}\n"
          f"В отчёте заказов: {len(res.orders)}\n"
          f"Удалено дубликатов: {res.duplicates}\n"
          f"Отброшено строк с ошибками: {res.rejected} (см. лист «Проблемы»)\n"
          f"Отчёт: {a.out}")
    return 0
