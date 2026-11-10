"""Чтение и очистка выгрузки заказов: разные форматы дат и чисел, дубликаты, битые строки."""
from __future__ import annotations

import csv
import datetime as dt
import io
import re
from collections import Counter
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path

FIELDS = ("order", "date", "client", "manager", "category", "qty", "price")
LABELS = {
    "order": "номер заказа", "date": "дата", "client": "клиент", "manager": "менеджер",
    "category": "категория", "qty": "количество", "price": "цена",
}
ALIASES = {
    "order": {"номер", "номер заказа", "заказ", "№", "№ заказа", "order", "order id", "id"},
    "date": {"дата", "дата заказа", "date"},
    "client": {"клиент", "покупатель", "контрагент", "client", "customer"},
    "manager": {"менеджер", "продавец", "manager", "seller"},
    "category": {"категория", "группа", "товарная группа", "category"},
    "qty": {"количество", "кол-во", "колво", "qty", "quantity"},
    "price": {"цена", "цена за единицу", "price"},
}
DATE_FORMATS = (
    "%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y",
    "%d.%m.%Y %H:%M", "%d.%m.%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S",
)
_UNIT_TAIL = re.compile(r"(?i)\s*(руб\.?|р\.?|₽|rub|шт\.?|кг)$")
_NUMBER = re.compile(r"^[+-]?\d+(\.\d+)?$")


def parse_number(value) -> Decimal:
    """Число из «1 234,50», «1,234.50», «2 500 ₽», 12, 3.5. Иначе ValueError."""
    if isinstance(value, bool):
        raise ValueError("не число")
    if isinstance(value, (int, float, Decimal)):
        d = Decimal(str(value))
        if not d.is_finite():
            raise ValueError("не число")
        return d
    s = str(value).strip().replace("−", "-").replace("–", "-")
    s = _UNIT_TAIL.sub("", s)
    s = re.sub(r"\s+", "", s)
    if not s:
        raise ValueError("пусто")
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif s.count(",") > 1:
        s = s.replace(",", "")
    elif s.count(".") > 1:
        s = s.replace(".", "")
    else:
        s = s.replace(",", ".")
    if not _NUMBER.match(s):
        raise ValueError("не число")
    try:
        return Decimal(s)
    except InvalidOperation:
        raise ValueError("не число")


def parse_date(value) -> dt.date:
    """Дата из строки (день идёт первым: 03.04.2026 — это 3 апреля) или из datetime/date."""
    if isinstance(value, dt.datetime):
        d = value.date()
    elif isinstance(value, dt.date):
        d = value
    else:
        s = str(value).strip()
        d = None
        for fmt in DATE_FORMATS:
            try:
                d = dt.datetime.strptime(s, fmt).date()
                break
            except ValueError:
                continue
        if d is None:
            raise ValueError("не дата")
    if not 2000 <= d.year <= 2100:
        raise ValueError("год вне диапазона 2000–2100")
    return d


def _norm_header(h) -> str:
    return re.sub(r"\s+", " ", str(h or "").replace("﻿", "")).strip().lower()


def _map_headers(headers) -> dict[int, str]:
    mapping: dict[int, str] = {}
    for i, h in enumerate(headers):
        n = _norm_header(h)
        for f, names in ALIASES.items():
            if n in names and f not in mapping.values():
                mapping[i] = f
                break
    missing = [LABELS[f] for f in FIELDS if f not in mapping.values()]
    if missing:
        raise ValueError("в файле нет колонок: " + ", ".join(missing))
    return mapping


def read_table(path) -> list[tuple[int, dict]]:
    """Читает .csv или .xlsx; возвращает [(номер строки в файле, {поле: значение})]."""
    p = Path(path)
    if p.suffix.lower() in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook
        wb = load_workbook(p, read_only=True, data_only=True)
        rows = list(wb.worksheets[0].iter_rows(values_only=True))
        wb.close()
    else:
        raw = p.read_bytes()
        try:
            text = raw.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = raw.decode("cp1251")
        first = text.split("\n", 1)[0]
        delim = max(";,\t", key=first.count)
        rows = list(csv.reader(io.StringIO(text), delimiter=delim))
    if not rows:
        raise ValueError("файл пустой")
    mapping = _map_headers(rows[0])
    out = []
    for n, row in enumerate(rows[1:], start=2):
        if all(c is None or str(c).strip() == "" for c in row):
            continue
        out.append((n, {f: (row[i] if i < len(row) else None) for i, f in mapping.items()}))
    return out


@dataclass
class Order:
    order: str
    date: dt.date
    client: str
    manager: str
    category: str
    qty: float
    price: float
    source_row: int

    @property
    def amount(self) -> float:
        return self.qty * self.price


@dataclass
class Issue:
    row: int
    kind: str          # «Отброшена» или «Дубликат»
    reason: str
    raw: str


@dataclass
class CleanResult:
    orders: list[Order] = field(default_factory=list)
    issues: list[Issue] = field(default_factory=list)
    total_rows: int = 0

    @property
    def duplicates(self) -> int:
        return sum(1 for i in self.issues if i.kind == "Дубликат")

    @property
    def rejected(self) -> int:
        return sum(1 for i in self.issues if i.kind == "Отброшена")


def _text(v) -> str:
    return re.sub(r"\s+", " ", "" if v is None else str(v)).strip()


def clean_rows(rows: list[tuple[int, dict]]) -> CleanResult:
    res = CleanResult(total_rows=len(rows))
    seen: dict[str, tuple] = {}
    for n, r in rows:
        raw = "; ".join(_text(r.get(f)) for f in FIELDS)
        reasons = []
        oid = _text(r.get("order"))
        if not oid:
            reasons.append("нет номера заказа")
        try:
            date = parse_date(r.get("date"))
        except ValueError as e:
            date = None
            reasons.append(f"дата: {e} («{_text(r.get('date'))}»)")
        texts = {}
        for f in ("client", "manager", "category"):
            texts[f] = _text(r.get(f))
            if not texts[f]:
                reasons.append(f"пустое поле: {LABELS[f]}")
        nums = {}
        for f, positive in (("qty", True), ("price", False)):
            try:
                v = parse_number(r.get(f))
                if (positive and v <= 0) or (not positive and v < 0):
                    raise ValueError("должно быть " + ("больше нуля" if positive else "не меньше нуля"))
                nums[f] = v
            except ValueError as e:
                reasons.append(f"{LABELS[f]}: {e} («{_text(r.get(f))}»)")
        if reasons:
            res.issues.append(Issue(n, "Отброшена", "; ".join(reasons), raw))
            continue
        key = oid.casefold()
        sig = (date, texts["client"].casefold(), texts["manager"].casefold(),
               texts["category"].casefold(), nums["qty"], nums["price"])
        if key in seen:
            first_row, first_sig = seen[key]
            if sig == first_sig:
                res.issues.append(Issue(n, "Дубликат", f"точная копия строки {first_row}", raw))
            else:
                res.issues.append(Issue(n, "Отброшена",
                                        f"номер {oid} уже есть в строке {first_row}, но данные другие", raw))
            continue
        seen[key] = (n, sig)
        res.orders.append(Order(oid, date, texts["client"], texts["manager"], texts["category"],
                                float(nums["qty"]), float(nums["price"]), n))
    _canonicalize(res.orders)
    return res


def _canonicalize(orders: list[Order]) -> None:
    """«иванов  И.», «Иванов И.» и «ИВАНОВ И.» — один человек: берём самое частое написание."""
    for f in ("client", "manager", "category"):
        spell: dict[str, Counter] = {}
        for o in orders:
            spell.setdefault(getattr(o, f).casefold(), Counter())[getattr(o, f)] += 1
        best = {k: sorted(c.items(), key=lambda kv: (-kv[1], kv[0]))[0][0] for k, c in spell.items()}
        for o in orders:
            setattr(o, f, best[getattr(o, f).casefold()])
