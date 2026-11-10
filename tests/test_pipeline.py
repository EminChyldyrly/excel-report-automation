import datetime as dt
import shutil
import subprocess
import tempfile
import unittest
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook, load_workbook

from xlreport.clean import clean_rows, read_table
from xlreport.cli import main
from xlreport.demo import generate
from xlreport.report import build_report, months_between

SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")


def recalc(path: Path, outdir: Path) -> Path:
    """Пересчитывает формулы в LibreOffice и возвращает путь к файлу со значениями."""
    subprocess.run([SOFFICE, "--headless", "--convert-to", "xlsx", "--outdir", str(outdir), str(path)],
                   check=True, capture_output=True, timeout=120)
    return outdir / path.name


class DemoRoundTrip(unittest.TestCase):
    """Очистка проверяется по «правде» генератора, а не по самой себе."""

    @classmethod
    def setUpClass(cls):
        cls.tmp = Path(tempfile.mkdtemp())
        cls.demo = generate(400, seed=3)
        (cls.tmp / "o.csv").write_text(cls.demo.csv_text, encoding="utf-8-sig")
        cls.res = clean_rows(read_table(cls.tmp / "o.csv"))

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def test_orders_equal_truth(self):
        got = {o.order: o for o in self.res.orders}
        self.assertEqual(set(got), {t["order"] for t in self.demo.truth})
        for t in self.demo.truth:
            o = got[t["order"]]
            self.assertEqual((o.date, o.client, o.manager, o.category), (t["date"], t["client"],
                                                                        t["manager"], t["category"]), t["order"])
            self.assertEqual((o.qty, o.price), (float(t["qty"]), t["price"]), t["order"])

    def test_issue_counts_match_generator(self):
        self.assertEqual(self.res.duplicates, self.demo.n_duplicates)
        self.assertEqual(self.res.rejected, self.demo.n_conflicts + self.demo.n_bad)
        self.assertEqual(self.res.total_rows, 400 + self.demo.n_duplicates + self.demo.n_conflicts + self.demo.n_bad)

    def test_deterministic(self):
        self.assertEqual(generate(50, 7).csv_text, generate(50, 7).csv_text)
        self.assertNotEqual(generate(50, 7).csv_text, generate(50, 8).csv_text)

    def test_report_structure(self):
        out = self.tmp / "r.xlsx"
        build_report(self.res, out)
        wb = load_workbook(out)
        self.assertEqual(wb.sheetnames, ["Сводка", "По месяцам", "По менеджерам", "По категориям",
                                         "Топ клиентов", "Данные", "Проблемы"])
        data = wb["Данные"]
        self.assertEqual(data.max_row, 1 + len(self.res.orders))
        self.assertEqual(data["I2"].value, "=G2*H2")                # сумма — формула
        self.assertEqual(wb["Проблемы"].max_row, 1 + len(self.res.issues))
        self.assertIn("SUMIFS", wb["По месяцам"]["B2"].value)
        self.assertEqual(len(wb["По месяцам"]._charts), 1)
        self.assertEqual(len(wb["Топ клиентов"]["B"]) - 1, 10)

    @unittest.skipUnless(SOFFICE, "LibreOffice не найден")
    def test_formulas_match_truth_after_recalculation(self):
        out = self.tmp / "calc.xlsx"
        build_report(self.res, out)
        outdir = self.tmp / "recalc"
        outdir.mkdir()
        wb = load_workbook(recalc(out, outdir), data_only=True)
        truth = self.demo.truth
        total = sum(t["qty"] * t["price"] for t in truth)
        s = wb["Сводка"]
        self.assertAlmostEqual(s["B3"].value, total, places=2)
        self.assertEqual(s["B4"].value, len(truth))
        self.assertAlmostEqual(s["B5"].value, total / len(truth), places=2)
        self.assertEqual(s["B6"].value.date(), min(t["date"] for t in truth))
        self.assertEqual(s["B7"].value.date(), max(t["date"] for t in truth))
        month = defaultdict(float)
        mgr = defaultdict(float)
        for t in truth:
            a = t["qty"] * t["price"]
            month[f"{t['date']:%Y-%m}"] += a
            mgr[t["manager"]] += a
        self.assertEqual(s["B8"].value, max(month, key=month.get))
        m = wb["По месяцам"]
        got = {m.cell(row=r, column=1).value: m.cell(row=r, column=2).value for r in range(2, m.max_row + 1)}
        self.assertEqual(set(got), set(months_between(min(t["date"] for t in truth),
                                                      max(t["date"] for t in truth))))
        for k, v in got.items():
            self.assertAlmostEqual(v, month.get(k, 0.0), places=2, msg=k)
        g = wb["По менеджерам"]
        for r in range(2, g.max_row + 1):
            self.assertAlmostEqual(g.cell(row=r, column=2).value, mgr[g.cell(row=r, column=1).value], places=2)
        revs = [g.cell(row=r, column=2).value for r in range(2, g.max_row + 1)]
        self.assertEqual(revs, sorted(revs, reverse=True))          # менеджеры — от большей выручки к меньшей
        c = wb["По категориям"]
        crevs = [c.cell(row=r, column=2).value for r in range(2, c.max_row + 1)]
        self.assertEqual(crevs, sorted(crevs, reverse=True))
        for r in range(3, m.max_row + 1):                            # динамика к предыдущему месяцу
            prev, cur = m.cell(row=r - 1, column=2).value, m.cell(row=r, column=2).value
            self.assertAlmostEqual(m.cell(row=r, column=5).value, cur / prev - 1, places=6)
        shares = sum(g.cell(row=r, column=4).value for r in range(2, g.max_row + 1))
        self.assertAlmostEqual(shares, 1.0, places=6)
        top = wb["Топ клиентов"]
        cl = defaultdict(float)
        for t in truth:
            cl[t["client"]] += t["qty"] * t["price"]
        want = sorted(cl.items(), key=lambda kv: (-kv[1], kv[0]))[:10]
        for i, (name, amount) in enumerate(want, start=2):
            self.assertEqual(top.cell(row=i, column=2).value, name)
            self.assertAlmostEqual(top.cell(row=i, column=3).value, amount, places=2)


class Edge(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_months_between(self):
        self.assertEqual(months_between(dt.date(2025, 11, 20), dt.date(2026, 2, 1)),
                         ["2025-11", "2025-12", "2026-01", "2026-02"])
        self.assertEqual(months_between(dt.date(2025, 5, 1), dt.date(2025, 5, 31)), ["2025-05"])

    def test_gap_months_listed_with_zero(self):
        csv = ("Номер;Дата;Клиент;Менеджер;Категория;Кол-во;Цена\n"
               "1;10.01.2025;К;М;Т;1;100\n2;05.04.2025;К;М;Т;2;50\n")
        (self.tmp / "g.csv").write_text(csv, encoding="utf-8")
        res = clean_rows(read_table(self.tmp / "g.csv"))
        out = self.tmp / "g.xlsx"
        build_report(res, out)
        ws = load_workbook(out)["По месяцам"]
        self.assertEqual([ws.cell(row=r, column=1).value for r in range(2, 6)],
                         ["2025-01", "2025-02", "2025-03", "2025-04"])

    def test_no_valid_rows_still_builds(self):
        (self.tmp / "e.csv").write_text("Номер;Дата;Клиент;Менеджер;Категория;Кол-во;Цена\n1;плохо;К;М;Т;1;1\n",
                                        encoding="utf-8")
        res = clean_rows(read_table(self.tmp / "e.csv"))
        out = self.tmp / "e.xlsx"
        build_report(res, out)
        self.assertEqual(load_workbook(out)["Проблемы"].max_row, 2)
        if SOFFICE:
            (self.tmp / "o").mkdir()
            s = load_workbook(recalc(out, self.tmp / "o"), data_only=True)["Сводка"]
            self.assertEqual(s["B3"].value, 0)
            self.assertEqual(s["B4"].value, 0)

    def test_missing_columns_message(self):
        (self.tmp / "m.csv").write_text("Номер;Дата\n1;01.01.2025\n", encoding="utf-8")
        with self.assertRaises(ValueError) as cm:
            read_table(self.tmp / "m.csv")
        self.assertIn("клиент", str(cm.exception))
        self.assertIn("цена", str(cm.exception))

    def test_cp1251_and_comma_delimiter(self):
        text = "Номер,Дата,Клиент,Менеджер,Категория,Количество,Цена\nA-1,05.03.2025,Ромашка,Иванов,Мебель,2,\"1 500,50\"\n"
        (self.tmp / "c.csv").write_bytes(text.encode("cp1251"))
        res = clean_rows(read_table(self.tmp / "c.csv"))
        self.assertEqual(len(res.orders), 1)
        self.assertEqual(res.orders[0].price, 1500.5)
        self.assertEqual(res.orders[0].client, "Ромашка")

    def test_xlsx_input_with_real_dates_and_numbers(self):
        wb = Workbook()
        ws = wb.active
        ws.append(["Заказ", "Дата заказа", "Покупатель", "Продавец", "Группа", "Qty", "Цена за единицу"])
        ws.append(["X1", dt.datetime(2025, 6, 1), "Ромашка", "Иванов", "Мебель", 3, 99.5])
        ws.append([None, None, None, None, None, None, None])           # пустая строка пропускается
        ws.append(["X2", "02.06.2025", "Ромашка", "Иванов", "Мебель", "2", "10,25"])
        wb.save(self.tmp / "i.xlsx")
        res = clean_rows(read_table(self.tmp / "i.xlsx"))
        self.assertEqual(res.total_rows, 2)
        self.assertEqual([(o.order, o.amount) for o in res.orders], [("X1", 298.5), ("X2", 20.5)])

    def test_cli_end_to_end(self):
        csv, xlsx = self.tmp / "o.csv", self.tmp / "r.xlsx"
        self.assertEqual(main(["demo", "-o", str(csv), "-n", "60", "--seed", "1"]), 0)
        self.assertEqual(main(["build", str(csv), "-o", str(xlsx), "--title", "Тест"]), 0)
        self.assertEqual(load_workbook(xlsx)["Сводка"]["A1"].value, "Тест")
        self.assertEqual(main(["build", str(self.tmp / "нет.csv"), "-o", str(xlsx)]), 2)


if __name__ == "__main__":
    unittest.main()
