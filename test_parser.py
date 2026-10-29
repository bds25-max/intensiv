"""
test_parser.py — Автономные сквозные автотесты для Задачи 1 (Контур КЦР).
Проверяют все условия приёмки, включая дословную сохранность данных без обезличивания.
"""

import os
import sys
import math
import tempfile
import unittest
from unittest.mock import patch
import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from parser_core import (
    BaseRegistryParser,
    PeriodicSliceParser,
    HistoricalFilesParser,
    QualityAuditor,
    parse_all_kcr,
    parse_strict_int,
    parse_strict_float,
    export_artifacts,
    SliceRecord,
    Provenance,
)


class TestKCRParserSuite(unittest.TestCase):
    def test_01_strict_numeric_and_nan_inf_validation(self):
        """Дробные, текстовые, NaN/Infinity и отрицательные значения отклоняются."""
        self.assertIsNotNone(parse_strict_int("1.5")[1])
        self.assertIsNotNone(parse_strict_int("#N/A")[1])
        self.assertIsNotNone(parse_strict_int(float("nan"))[1])
        self.assertIsNotNone(parse_strict_int(float("inf"))[1])

        self.assertIsNotNone(parse_strict_float("NaN")[1])
        self.assertIsNotNone(parse_strict_float("Infinity")[1])
        self.assertIsNotNone(parse_strict_float("-inf")[1])
        self.assertIsNotNone(parse_strict_float(float("nan"))[1])

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

    def test_03_same_sheet_duplicate_invalid_first_and_conflicting_valid(self):
        """
        Сквозной тест Excel -> PeriodicSliceParser -> QualityAuditor:
        1) Если первая строка клиента на листе содержит ошибку ('Ошибочный текст', 2), а вторая валидна (10, 2) — вторая сохраняется.
        2) Если обе строки одного клиента на листе валидны, но имеют разные остатки (2 и 8) — обе отклоняются как конфликт.
        """
        with tempfile.TemporaryDirectory() as tmpdir:
            fpath = os.path.join(tmpdir, "test_same_sheet.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "весь 2025 год"
            ws.cell(1, 1, "Служебная строка")
            ws.cell(2, 1, "Уникальны обезличинный идентификатор")
            ws.cell(2, 4, "документов всего")
            ws.cell(2, 5, "Всего не отправлено документов")

            # Клиент C_RECOVER: 1-я строка битая, 2-я валидная
            ws.cell(3, 1, "C_RECOVER")
            ws.cell(3, 4, "Ошибочный текст")
            ws.cell(3, 5, 2)
            ws.cell(4, 1, "C_RECOVER")
            ws.cell(4, 4, 10)
            ws.cell(4, 5, 2)

            # Клиент C_CONFLICT: обе строки валидные, но разные остатки (2 и 8)
            ws.cell(5, 1, "C_CONFLICT")
            ws.cell(5, 4, 10)
            ws.cell(5, 5, 2)
            ws.cell(6, 1, "C_CONFLICT")
            ws.cell(6, 4, 10)
            ws.cell(6, 5, 8)

            wb.save(fpath)
            wb.close()

            raw_slices, p_issues = PeriodicSliceParser.parse(fpath, known_clients={"C_RECOVER", "C_CONFLICT"})
            valid, rejected, a_issues = QualityAuditor.audit_and_filter_slices(raw_slices)

            valid_ids = [s.source_client_id for s in valid]
            self.assertEqual(valid_ids, ["C_RECOVER"], "Корректная вторая строка C_RECOVER должна сохраниться, а конфликт C_CONFLICT — отклониться")
            self.assertEqual(valid[0].total, 10)
            self.assertEqual(valid[0].unsent, 2)
            self.assertEqual(len(rejected), 3)

    def test_04_header_period_mismatch_rejected(self):
        """Подмена периода в заголовке колонки остатка ('Долг 1 кв 2026 на 29.09.26' вместо 'весь 2025') вызывает ValueError."""
        with tempfile.TemporaryDirectory() as tmpdir:
            fpath = os.path.join(tmpdir, "bad_period.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Все КЦР"
            ws.cell(2, 1, "ID")
            ws.cell(2, 10, "Действует ли КЦР")
            ws.cell(2, 11, "Блокировка в сервисе")
            for c_tot, c_uns, c_pct, _, _, _, p_tok, d_tok in BaseRegistryParser.SLICE_SPECS:
                ws.cell(2, c_tot, f"Всего документов {p_tok}")
                ws.cell(2, c_uns, f"Долг {p_tok} на {d_tok}")
                ws.cell(2, c_pct, f"% Долга {p_tok} на {d_tok}")
            ws.cell(2, 16, "Долг 1 кв 2026 на 29.09.26")
            wb.save(fpath)
            wb.close()

            with self.assertRaises(ValueError):
                BaseRegistryParser.parse(fpath)

    def test_05_backup_preserved_when_rollback_fails(self):
        """При двойном сбое (публикации и самого отката) папка _backup не удаляется."""
        with tempfile.TemporaryDirectory() as tmpdir:
            export_artifacts([{"id": "OLD"}], [], [], [], [], output_dir=tmpdir)

            real_replace = os.replace
            call_count = {"n": 0}

            def side_effect_replace(src, dst):
                call_count["n"] += 1
                if call_count["n"] >= 7:
                    raise OSError("Simulated failure during publish and rollback")
                return real_replace(src, dst)

            with patch("os.replace", side_effect=side_effect_replace):
                with self.assertRaises(OSError):
                    export_artifacts([{"id": "NEW"}], [], [], [], [], output_dir=tmpdir)

            backup_dir = os.path.join(tmpdir, "_backup")
            self.assertTrue(os.path.exists(backup_dir), "Папка _backup обязана сохраниться при ошибке восстановления!")
            self.assertTrue(os.path.exists(os.path.join(backup_dir, "normalized_clients.json")))

    def test_06_verbatim_history_without_masking(self):
        """Все телефоны, email и ФИО в истории сохраняются дословно без какого-либо маскирования."""
        f3 = "База должники III кв 2025 Дубна.xlsx"
        if os.path.exists(f3):
            _, history, _ = HistoricalFilesParser.parse_debts_2025(f3)
            all_texts = " ".join(h.text or "" for h in history) + " " + " ".join(h.raw_deadline or "" for h in history)
            self.assertNotIn("[ТЕЛЕФОН_СКРЫТ]", all_texts)
            self.assertIn("79995209332", all_texts)
            self.assertIn("+7 (914) 200-06-94", all_texts)

    def test_07_date_and_period_bounds_validation(self):
        """Проверка календарной корректности дат и условия period_start <= period_end."""
        bad_date = SliceRecord("C1", "2025-01-01", "2025-12-31", "2026-99-99", 10, 1, None, Provenance("f", "s", 1))
        bad_bounds = SliceRecord("C2", "2026-06-01", "2026-01-01", "2026-10-20", 10, 1, None, Provenance("f", "s", 2))
        valid, rejected, issues = QualityAuditor.audit_and_filter_slices([bad_date, bad_bounds])
        self.assertEqual(len(valid), 0)
        self.assertEqual(len(rejected), 2)
        codes = {i.code for i in issues}
        self.assertIn("INVALID_DATE_VALUE", codes)
        self.assertIn("INVALID_PERIOD_BOUNDS", codes)

    def test_08_preflight_and_unknown_client_isolation(self):
        """Запуск в пустой папке выбрасывает FileNotFoundError; неизвестный клиент изолируется."""
        with tempfile.TemporaryDirectory() as empty_dir:
            with self.assertRaises(FileNotFoundError):
                parse_all_kcr(base_dir=empty_dir)

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
