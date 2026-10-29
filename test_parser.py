"""
test_parser.py — Автономные автотесты для Задачи 1 (Контур КЦР).
Проверяют все условия приёмки, включая регресс-кейсы повторного ревью.
"""

import os
import sys
import tempfile
import unittest
import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from parser_core import (
    BaseRegistryParser,
    PeriodicSliceParser,
    HistoricalFilesParser,
    QualityAuditor,
    parse_all_kcr,
    parse_strict_int,
    SliceRecord,
    Provenance,
)


class TestKCRParserSuite(unittest.TestCase):
    def test_01_strict_int_and_negative_validation(self):
        """Дробные, текстовые и отрицательные количества отклоняются без молчаливых округлений."""
        val, err = parse_strict_int("1.5")
        self.assertIsNotNone(err)
        self.assertIsNone(val)

        val, err = parse_strict_int("#N/A")
        self.assertIsNotNone(err)
        self.assertIsNone(val)

        # Отрицательные числа должны отклоняться валидатором
        s1 = SliceRecord("C1", "2025-01-01", "2025-12-31", "2026-10-20", 10, -2, None, Provenance("f", "s", 1))
        s2 = SliceRecord("C2", "2025-01-01", "2025-12-31", "2026-10-20", -1, -2, None, Provenance("f", "s", 2))
        valid, rejected, issues = QualityAuditor.audit_and_filter_slices([s1, s2])
        self.assertEqual(len(valid), 0)
        self.assertEqual(len(rejected), 2)
        codes = {i.code for i in issues}
        self.assertIn("NEGATIVE_UNSENT", codes)
        self.assertIn("NEGATIVE_TOTAL", codes)

    def test_02_invalid_first_record_does_not_block_valid_second(self):
        """Ошибочная первая запись (total=None, unsent=0) не блокирует следующую валидную запись (total=0, unsent=0)."""
        bad_first = SliceRecord(
            "XXD78058", "2025-04-01", "2026-03-31", "2026-05-21",
            total=None, unsent=0, raw_percentage=None,
            provenance=Provenance("КЦР_последняя выгрузка_Дубна.xlsx", "Все КЦР", 5004, "AN,AS")
        )
        good_second = SliceRecord(
            "XXD78058", "2025-04-01", "2026-03-31", "2026-05-21",
            total=0, unsent=0, raw_percentage=None,
            provenance=Provenance("База I кв 2026 Дубна.xlsx", "Все КЦР", 98, "AE,AF")
        )
        valid, rejected, issues = QualityAuditor.audit_and_filter_slices([bad_first, good_second])
        self.assertEqual(len(valid), 1, "Вторая валидная запись обязана пройти в valid")
        self.assertEqual(valid[0].total, 0)
        self.assertEqual(len(rejected), 1)
        self.assertEqual(rejected[0].rejection_reason, "UNSENT_WITHOUT_TOTAL")

    def test_03_date_and_period_bounds_validation(self):
        """Проверка календарной корректности дат и условия period_start <= period_end."""
        bad_date = SliceRecord("C1", "2025-01-01", "2025-12-31", "2026-99-99", 10, 1, None, Provenance("f", "s", 1))
        bad_bounds = SliceRecord("C2", "2026-06-01", "2026-01-01", "2026-10-20", 10, 1, None, Provenance("f", "s", 2))
        valid, rejected, issues = QualityAuditor.audit_and_filter_slices([bad_date, bad_bounds])
        self.assertEqual(len(valid), 0)
        self.assertEqual(len(rejected), 2)
        codes = {i.code for i in issues}
        self.assertIn("INVALID_DATE_VALUE", codes)
        self.assertIn("INVALID_PERIOD_BOUNDS", codes)

    def test_04_preflight_missing_files_fails(self):
        """Запуск в пустой папке выбрасывает FileNotFoundError."""
        with tempfile.TemporaryDirectory() as empty_dir:
            with self.assertRaises(FileNotFoundError):
                parse_all_kcr(base_dir=empty_dir)

    def test_05_unknown_client_physical_isolation(self):
        """Неизвестный клиент изолируется в rejected_slices и не попадает в valid."""
        with tempfile.TemporaryDirectory() as tmpdir:
            fpath = os.path.join(tmpdir, "test_slice.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "весь 2025 год"
            ws.cell(1, 1, "Служебная строка")
            ws.cell(2, 1, "Уникальны обезличинный идентификатор")
            ws.cell(2, 4, "документов всего")
            ws.cell(2, 5, "Всего не отправлено документов")
            ws.cell(3, 1, "UNKNOWN_CLIENT_999")
            ws.cell(3, 4, 100)
            ws.cell(3, 5, 10)
            wb.save(fpath)
            wb.close()

            slices, _ = PeriodicSliceParser.parse(fpath, known_clients={"CLIENT_AAA"})
            valid, rejected, _ = QualityAuditor.audit_and_filter_slices(slices)
            self.assertEqual(len(valid), 0)
            self.assertEqual(len(rejected), 1)
            self.assertEqual(rejected[0].rejection_reason, "UNKNOWN_CLIENT_ID")


if __name__ == "__main__":
    unittest.main()
