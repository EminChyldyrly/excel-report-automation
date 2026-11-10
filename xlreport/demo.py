"""Генератор «грязной» выгрузки заказов с известной правдой — для демо и для проверки очистки."""
from __future__ import annotations

import datetime as dt
import random
from dataclasses import dataclass

CLIENTS = ["ООО Ромашка", "ИП Петров А.В.", "ООО Вектор", "АО Северстрой", "ИП Сидорова Н.К.",
           "ООО ТехноПарк", "ЗАО Альфа-Опт", "ООО Лидер", "ИП Кузнецов Д.С.", "ООО Горизонт",
           "АО Меридиан", "ООО Прогресс"]
MANAGERS = ["Иванова Анна", "Смирнов Павел", "Орлова Мария", "Фёдоров Илья"]
CATEGORIES = {"Канцелярия": (40, 900), "Офисная техника": (3500, 45000), "Мебель": (6000, 60000),
              "Расходные материалы": (300, 7000), "Упаковка": (80, 2500)}


@dataclass
class Demo:
    csv_text: str
    truth: list[dict]      # правильные заказы, какими они должны получиться после очистки
    n_duplicates: int
    n_conflicts: int
    n_bad: int


def _fmt_date(rng, d: dt.date) -> str:
    return d.strftime(rng.choice(["%d.%m.%Y", "%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%d.%m.%y"]))


def _fmt_num(rng, x: float, money: bool) -> str:
    style = rng.randrange(4)
    if not money:
        return str(int(x)) if style < 3 else f"{int(x)} шт"
    whole = f"{x:.2f}"
    if style == 0:
        return whole
    if style == 1:
        return whole.replace(".", ",")
    if style == 2:
        a, b = whole.split(".")
        return f"{int(a):,}".replace(",", " ") + "," + b
    a, b = whole.split(".")
    return f"{int(a):,}".replace(",", " ") + "," + b + " ₽"


def _mess(rng, s: str) -> str:
    r = rng.random()
    if r < 0.08:
        return s.lower()
    if r < 0.14:
        return s.upper()
    if r < 0.2:
        return "  " + s.replace(" ", "  ") + " "
    return s


def generate(n: int = 500, seed: int = 0, year: int = 2025) -> Demo:
    rng = random.Random(seed)
    start = dt.date(year, 1, 1)
    truth, lines = [], []
    for i in range(n):
        cat = rng.choice(list(CATEGORIES))
        lo, hi = CATEGORIES[cat]
        d = start + dt.timedelta(days=rng.randrange(365))
        # весной и осенью заказов больше
        if rng.random() < 0.3:
            d = dt.date(year, rng.choice([3, 4, 9, 10]), rng.randrange(1, 28))
        o = {"order": f"A-{1000 + i}", "date": d, "client": rng.choice(CLIENTS),
             "manager": rng.choice(MANAGERS), "category": cat,
             "qty": rng.randrange(1, 40), "price": round(rng.uniform(lo, hi), 2)}
        truth.append(o)
    rows = []
    for o in truth:
        rows.append([o["order"], _fmt_date(rng, o["date"]), _mess(rng, o["client"]),
                     _mess(rng, o["manager"]), _mess(rng, o["category"]),
                     _fmt_num(rng, o["qty"], False), _fmt_num(rng, o["price"], True)])
    n_dup = n_conf = n_bad = 0
    after: dict[int, list[list[str]]] = {}                    # копии ставим строго после оригинала
    for i in rng.sample(range(len(rows)), k=max(1, n // 30)):  # точные дубликаты
        after.setdefault(i, []).append(list(rows[i]))
        n_dup += 1
    for i in rng.sample(range(len(rows)), k=max(1, n // 100)):  # тот же номер, другие данные
        r = list(rows[i])
        r[5] = str(int(float(r[5].split()[0])) + 1)
        after.setdefault(i, []).append(r)
        n_conf += 1
    bad_kinds = [
        lambda r: r.__setitem__(1, "31.02.2025"),
        lambda r: r.__setitem__(5, "-3"),
        lambda r: r.__setitem__(2, ""),
        lambda r: r.__setitem__(6, "договорная"),
        lambda r: r.__setitem__(0, ""),
        lambda r: r.__setitem__(1, "вчера"),
    ]
    allrows: list[list[str]] = []
    for i, r in enumerate(rows):
        allrows.append(r)
        allrows.extend(after.get(i, []))
    for k in range(max(2, n // 25)):
        r = [f"B-{k}", "05.05.2025", "ООО Лидер", "Орлова Мария", "Мебель", "2", "1500,00"]
        bad_kinds[k % len(bad_kinds)](r)
        allrows.insert(rng.randrange(len(allrows) + 1), r)
        n_bad += 1
    head = ["Номер заказа", "Дата", "Клиент", "Менеджер", "Категория", "Количество", "Цена"]
    text = "\n".join(";".join(c if ";" not in c else f'"{c}"' for c in r) for r in [head] + allrows) + "\n"
    return Demo(text, truth, n_dup, n_conf, n_bad)
