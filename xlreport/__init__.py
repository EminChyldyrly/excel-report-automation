"""xlreport: из выгрузки заказов (CSV/XLSX) делает Excel-отчёт с формулами и графиками."""
from .clean import clean_rows, parse_date, parse_number, read_table
from .report import build_report

__all__ = ["clean_rows", "parse_date", "parse_number", "read_table", "build_report"]
__version__ = "0.1.0"
