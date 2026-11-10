import datetime as dt
import unittest
from decimal import Decimal

from xlreport.clean import clean_rows, parse_date, parse_number


def row(n, order="A-1", date="05.03.2025", client="ООО Лидер", manager="Орлова Мария",
        category="Мебель", qty="2", price="1500"):
    return n, dict(order=order, date=date, client=client, manager=manager,
                   category=category, qty=qty, price=price)


class ParseNumber(unittest.TestCase):
    def test_good(self):
        cases = {"1234": "1234", "1 234,50": "1234.50", "1,234.50": "1234.50", "1.234,50": "1234.50",
                 "2 500 ₽": "2500", "2500 руб.": "2500", "12,5": "12.5", "3 шт": "3",
                 "−7": "-7", "1,234,567": "1234567", "0": "0", " 42 ": "42"}
        for src, want in cases.items():
            self.assertEqual(parse_number(src), Decimal(want), src)
        self.assertEqual(parse_number(12), Decimal(12))
        self.assertEqual(parse_number(3.5), Decimal("3.5"))

    def test_bad(self):
        for src in ["", "abc", "договорная", "1.2.3,4x", "--5", None, True, float("nan"), "12 abc"]:
            with self.assertRaises(ValueError, msg=repr(src)):
                parse_number(src)


class ParseDate(unittest.TestCase):
    def test_good(self):
        want = dt.date(2025, 3, 5)
        for s in ["05.03.2025", "5.3.2025", "2025-03-05", "05/03/2025", "05.03.25", "05-03-2025",
                  "05.03.2025 14:30", "2025-03-05 10:00:00"]:
            self.assertEqual(parse_date(s), want, s)
        self.assertEqual(parse_date(dt.datetime(2025, 3, 5, 9, 0)), want)
        self.assertEqual(parse_date(want), want)

    def test_day_comes_first(self):
        self.assertEqual(parse_date("03.04.2025"), dt.date(2025, 4, 3))

    def test_year_bounds(self):
        self.assertEqual(parse_date("01.01.2000"), dt.date(2000, 1, 1))
        self.assertEqual(parse_date("31.12.2100"), dt.date(2100, 12, 31))
        for s in ["31.12.1999", "01.01.2101"]:
            with self.assertRaises(ValueError, msg=s):
                parse_date(s)

    def test_bad(self):
        for s in ["31.02.2025", "вчера", "", "13.13.2025", "01.01.1850", "2025-02-30"]:
            with self.assertRaises(ValueError, msg=s):
                parse_date(s)


class CleanRows(unittest.TestCase):
    def test_valid_row(self):
        r = clean_rows([row(2)])
        self.assertEqual(len(r.orders), 1)
        o = r.orders[0]
        self.assertEqual((o.date, o.qty, o.price, o.amount), (dt.date(2025, 3, 5), 2.0, 1500.0, 3000.0))
        self.assertEqual(r.issues, [])

    def test_exact_duplicate_dropped_once(self):
        r = clean_rows([row(2), row(3), row(4, client="ооо лидер")])
        self.assertEqual(len(r.orders), 1)
        self.assertEqual([(i.row, i.kind) for i in r.issues], [(3, "Дубликат"), (4, "Дубликат")])

    def test_conflicting_duplicate_rejected(self):
        r = clean_rows([row(2), row(3, qty="5")])
        self.assertEqual(len(r.orders), 1)
        self.assertEqual(r.orders[0].qty, 2.0)
        self.assertEqual(r.issues[0].kind, "Отброшена")
        self.assertIn("строке 2", r.issues[0].reason)

    def test_bad_rows_have_reasons(self):
        rows = [row(2, date="31.02.2025"), row(3, qty="-3"), row(4, client=""), row(5, price="договорная"),
                row(6, order=""), row(7, qty="0"), row(8, price="-1")]
        r = clean_rows(rows)
        self.assertEqual(r.orders, [])
        self.assertEqual(r.rejected, 7)
        text = " | ".join(i.reason for i in r.issues)
        for part in ["дата", "количество", "клиент", "цена", "номера заказа"]:
            self.assertIn(part, text)

    def test_bad_first_row_does_not_block_valid_second(self):
        r = clean_rows([row(2, qty="-1"), row(3, qty="4")])
        self.assertEqual([o.qty for o in r.orders], [4.0])

    def test_zero_price_allowed(self):
        self.assertEqual(len(clean_rows([row(2, price="0")]).orders), 1)

    def test_names_unified_to_most_common_spelling(self):
        rows = [row(2, "A-1", client="ООО Лидер"), row(3, "A-2", client="ооо лидер"),
                row(4, "A-3", client="  ООО   Лидер "), row(5, "A-4", client="ООО ЛИДЕР")]
        r = clean_rows(rows)
        self.assertEqual({o.client for o in r.orders}, {"ООО Лидер"})

    def test_order_ids_compared_ignoring_case_and_spaces(self):
        r = clean_rows([row(2, order="a-1"), row(3, order=" A-1 ")])
        self.assertEqual(len(r.orders), 1)
        self.assertEqual(r.duplicates, 1)

    def test_counts(self):
        r = clean_rows([row(2), row(3), row(4, "A-2", qty="x"), row(5, "A-3")])
        self.assertEqual((r.total_rows, len(r.orders), r.duplicates, r.rejected), (4, 2, 1, 1))


if __name__ == "__main__":
    unittest.main()
