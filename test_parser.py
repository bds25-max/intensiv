"""
test_parser.py — Автономные автотесты для Задачи 1 (Контур КЦР).
Совместимы с unittest и pytest.
"""

import os
import sys
import tempfile
import unittest
import openpyxl

# Обеспечиваем импорт parser_core из текущей директории файла
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from parser_core import (
    BaseRegistryParser,
    PeriodicSliceParser,
    HistoricalFilesParser,
    QualityAuditor,
    parse_all_kcr,
    parse_strict_int,
    SliceRecord,
    Provenance
)

class TestKCRParserSuite(unittest.TestCase):
    def test_01_strict_int_validation(self):
        """Проверка строгости чисел: дробные и битые строки не округляются молча."""
        val, err = parse_strict_int("1.5")
        self.assertIsNotNone(err, "Дробное число должно возвращать ошибку, а не округляться")
        self.assertIsNone(val)

        val, err = parse_strict_int("abc")
        self.assertIsNotNone(err)
        self.assertIsNone(val)

        val, err = parse_strict_int("#VALUE!")
        self.assertIsNotNone(err)
        self.assertIsNone(val)

        val, err = parse_strict_int(None)
        self.assertIsNone(err)
        self.assertIsNone(val, "None должен оставаться None")

        val, err = parse_strict_int(25)
        self.assertIsNone(err)
        self.assertEqual(val, 25)

    def test_02_preflight_missing_files_fails(self):
        """Проверка: запуск в пустой папке должен выбросить FileNotFoundError, а не завершаться 0."""
        with tempfile.TemporaryDirectory() as empty_dir:
            with self.assertRaises(FileNotFoundError):
                parse_all_kcr(base_dir=empty_dir)

    def test_03_unknown_client_physical_isolation(self):
        """Проверка: неизвестный клиент получает статус blocked и изолируется в rejected_slices."""
        with tempfile.TemporaryDirectory() as tmpdir:
            fpath = os.path.join(tmpdir, "test_slice.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "весь 2025 год"
            ws.cell(1, 1, "Служебная строка")
            ws.cell(2, 1, "Уникальны обезличинный идентификатор")
            ws.cell(2, 4, "всего")
            ws.cell(2, 5, "не отправлено")
            ws.cell(3, 1, "UNKNOWN_CLIENT_999")
            ws.cell(3, 4, 100)
            ws.cell(3, 5, 10)
            wb.save(fpath)

            slices, issues = PeriodicSliceParser.parse(fpath, known_clients={"CLIENT_AAA"})
            self.assertEqual(len(slices), 1)
            self.assertEqual(slices[0].rejection_reason, "UNKNOWN_CLIENT_ID")

            valid, rejected, _ = QualityAuditor.audit_and_filter_slices(slices)
            self.assertEqual(len(valid), 0, "Неизвестный клиент не должен попадать в валидные срезы")
            self.assertEqual(len(rejected), 1, "Неизвестный клиент обязан попасть в rejected_slices")

    def test_04_duplicate_slice_quarantine(self):
        """Проверка: дублирующиеся срезы одного ключа попадают в карантин (rejected)."""
        with tempfile.TemporaryDirectory() as tmpdir:
            fpath = os.path.join(tmpdir, "test_slice.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "1 кв 2026"
            ws.cell(2, 1, "ID")
            ws.cell(2, 4, "всего")
            ws.cell(2, 5, "не отправлено")
            ws.cell(3, 1, "CLIENT_DUP")
            ws.cell(3, 4, 50)
            ws.cell(3, 5, 5)
            ws.cell(4, 1, "CLIENT_DUP")
            ws.cell(4, 4, 50)
            ws.cell(4, 5, 5)
            wb.save(fpath)

            slices, issues = PeriodicSliceParser.parse(fpath, known_clients={"CLIENT_DUP"})
            valid, rejected, audit_issues = QualityAuditor.audit_and_filter_slices(slices)
            self.assertEqual(len(valid), 1)
            self.assertEqual(len(rejected), 1, "Дубль среза должен изолироваться в rejected_slices")

    def test_05_unsent_without_total_flagged(self):
        """Проверка: total=None при заполненном остатке регистрируется как ошибка."""
        record = SliceRecord(
            source_client_id="C1", period_start="2025-01-01", period_end="2025-12-31",
            as_of_date="2026-10-20", total=None, unsent=15, raw_percentage=None,
            provenance=Provenance("f", "s", 1)
        )
        valid, rejected, issues = QualityAuditor.audit_and_filter_slices([record])
        self.assertEqual(len(valid), 0)
        self.assertEqual(len(rejected), 1)
        codes = {i.code for i in issues}
        self.assertIn("UNSENT_WITHOUT_TOTAL", codes)

if __name__ == "__main__":
    unittest.main()
