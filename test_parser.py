"""
test_parser.py — Автономные сквозные автотесты для Задачи 1 (Контур КЦР).
Тесты выполняют проверочные утверждения (assertions)
даже при запуске в изолированной папке без исходных Excel-таблиц.
"""

import os
import copy
import itertools
import sys
import math
import json
import tempfile
import unittest
from unittest.mock import patch
from dataclasses import asdict
import openpyxl

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

from parser_core import (
    BaseRegistryParser,
    PeriodicSliceParser,
    HistoricalFilesParser,
    QualityAuditor,
    parse_all_kcr,
    parse_strict_int,
    parse_strict_float,
    export_artifacts,
    verify_header_exact_period,
    extract_period_from_header,
    COLUMN_MAPPING,
    SliceRecord,
    ClientRecord,
    Provenance,
)


def locate_excel_file(filename: str):
    for d in [os.path.join(SCRIPT_DIR, "sources"), ".", "..", "../..", SCRIPT_DIR, os.path.join(SCRIPT_DIR, ".."), os.path.join(SCRIPT_DIR, "../..")]:
        p = os.path.join(d, filename)
        if os.path.exists(p):
            return p
    return None


class TestKCRParserSuite(unittest.TestCase):
    def test_01_strict_numeric_and_nan_inf_validation(self):
        """Дробные (вкл. 2.0000000000000001), текстовые, формульные, NaN/Infinity и отрицательные значения отклоняются."""
        self.assertIsNotNone(parse_strict_int("1.5")[1])
        self.assertIsNotNone(parse_strict_int("2.0000000000000001")[1])
        self.assertEqual(parse_strict_int("2.0")[0], 2)
        self.assertIsNotNone(parse_strict_int("#N/A")[1])
        self.assertIsNotNone(parse_strict_int("=2+2")[1])
        self.assertIsNotNone(parse_strict_int(float("nan"))[1])
        self.assertIsNotNone(parse_strict_int(float("inf"))[1])

        self.assertIsNotNone(parse_strict_float("NaN")[1])
        self.assertIsNotNone(parse_strict_float("Infinity")[1])
        self.assertIsNotNone(parse_strict_float("=B2/C2")[1])
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
        """Ошибочная первая запись (total=None, unsent=0) на ту же дату не блокирует следующую валидную запись (total=0, unsent=0)."""
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
        valid, rejected, _ = QualityAuditor.audit_and_filter_slices([bad_first, good_second])
        self.assertEqual(len(valid), 1, "Вторая валидная запись обязана пройти в valid")
        self.assertEqual(valid[0].total, 0)
        self.assertEqual(len(rejected), 1)
        self.assertEqual(rejected[0].rejection_reason, "UNSENT_WITHOUT_TOTAL")

    def test_03_same_sheet_duplicate_recovery_and_conflict(self):
        """Сквозной тест на листе: битая 1-я строка не блокирует 2-ю валидную; две разные валидные дают конфликт."""
        with tempfile.TemporaryDirectory() as tmpdir:
            fpath = os.path.join(tmpdir, "test_same_sheet.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "весь 2025 год"
            ws.cell(1, 1, "Служебная строка")
            ws.cell(2, 1, "Уникальны обезличинный идентификатор")
            ws.cell(2, 4, "документов всего")
            ws.cell(2, 5, "Всего не отправлено документов")

            ws.cell(3, 1, "C_REC")
            ws.cell(3, 4, "Ошибочный текст")
            ws.cell(3, 5, 2)
            ws.cell(4, 1, "C_REC")
            ws.cell(4, 4, 10)
            ws.cell(4, 5, 2)

            ws.cell(5, 1, "C_CONF")
            ws.cell(5, 4, 10)
            ws.cell(5, 5, 2)
            ws.cell(6, 1, "C_CONF")
            ws.cell(6, 4, 10)
            ws.cell(6, 5, 8)

            wb.save(fpath)
            wb.close()

            raw_slices, _ = PeriodicSliceParser.parse(fpath, known_clients={"C_REC", "C_CONF"})
            valid, rejected, _ = QualityAuditor.audit_and_filter_slices(raw_slices)

            valid_ids = [s.source_client_id for s in valid]
            self.assertEqual(valid_ids, ["C_REC"])
            self.assertEqual(valid[0].total, 10)
            self.assertEqual(len(rejected), 3)

    def test_04_exact_period_bounds_verification(self):
        """Полное сравнение периода отвергает любые накопительные интервалы ('весь 2025 — II квартал 2026', 'весь 2025 год - 2 кв 2026')."""
        self.assertEqual(extract_period_from_header("Долг весь 2025 на 29.09.26"), "весь 2025")
        self.assertEqual(extract_period_from_header("Долг весь 2025 — II квартал 2026 на 29.09.26"), "весь 2025-2 кв 2026")
        self.assertEqual(extract_period_from_header("Долг весь 2025 год - 2 кв 2026 на 29.09.26"), "весь 2025-2 кв 2026")

        self.assertTrue(verify_header_exact_period("Долг весь 2025 на 29.09.26", "весь 2025"))
        self.assertFalse(verify_header_exact_period("Долг весь 2025 - 2 кв 2026 на 29.09.26", "весь 2025"))
        self.assertFalse(verify_header_exact_period("Долг весь 2025 — II квартал 2026 на 29.09.26", "весь 2025"))
        self.assertFalse(verify_header_exact_period("Долг весь 2025 год — II квартал 2026 на 29.09.26", "весь 2025"))
        self.assertFalse(verify_header_exact_period("Документов всего за 1-3 кв 2025 год", "весь 2025"))
        self.assertTrue(verify_header_exact_period("Документов всего за 2025 год", "весь 2025"))
        self.assertTrue(verify_header_exact_period("Долг весь 2025 - 2 кв 2026 на 29.09.26", "весь 2025-2 кв 2026"))
        self.assertTrue(verify_header_exact_period("Долг весь 2025 — II квартал 2026 на 29.09.26", "весь 2025-2 кв 2026"))

        # Сквозная проверка на книге Excel: подмена 'весь 2025' на 'весь 2025 — II квартал 2026' вызывает ValueError
        with tempfile.TemporaryDirectory() as tmpdir:
            fpath = os.path.join(tmpdir, "bad_cum_period.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Все КЦР"
            ws.cell(2, 1, "ID")
            ws.cell(2, 6, "Статус на 21.07.26")
            ws.cell(2, 7, "стаус на 21.05.26")
            ws.cell(2, 10, "Действует ли КЦР")
            ws.cell(2, 11, "Блокировка в сервисе")
            for c_tot, c_uns, c_pct, _, _, _, p_tok, d_tok in BaseRegistryParser.SLICE_SPECS:
                ws.cell(2, c_tot, f"Всего документов {p_tok}")
                ws.cell(2, c_uns, f"Долг {p_tok} на {d_tok}")
                ws.cell(2, c_pct, f"% Долга {p_tok} на {d_tok}")
            ws.cell(2, 16, "Долг весь 2025 — II квартал 2026 на 29.09.26")
            wb.save(fpath)
            wb.close()

            with self.assertRaises(ValueError):
                BaseRegistryParser.parse(fpath)

    def test_05_client_model_rich_attributes_saved(self):
        """Поля ИНН, КПП, телефон, email, код КЦР и СЦ корректно сохраняются в карточке клиента."""
        c = ClientRecord(
            source_client_id="TEST_CID",
            is_active="Дейсвующий",
            is_blocked="нет",
            time_zone_offset="3",
            admin_name="Иванов И.И.",
            provenance=Provenance("f", "s", 1),
            kcr_code="KCR-1234",
            organization_name="ООО Контур-Тест",
            service_center="СЦ Восток",
            inn="7701234567",
            kpp="770101001",
            creation_date="2024-01-15",
            admin_phone="+7 (999) 111-22-33",
            admin_email="admin@test.ru",
            manager_name="Петров П.П."
        )
        d = asdict(c)
        self.assertEqual(d["kcr_code"], "KCR-1234")
        self.assertEqual(d["inn"], "7701234567")
        self.assertEqual(d["admin_phone"], "+7 (999) 111-22-33")
        self.assertEqual(d["manager_name"], "Петров П.П.")

    def test_06_manager_separated_from_admin(self):
        """Имя менеджера из выгрузки не затирает имя главного администратора, а provenance обогащения сохраняется."""
        clients_dict = {
            "CID1": ClientRecord(
                source_client_id="CID1",
                is_active="Дейсвующий",
                is_blocked="нет",
                time_zone_offset=None,
                admin_name=None,
                manager_name="Менеджер Сидоров",
                provenance=Provenance("f1", "s1", 1)
            )
        }
        with tempfile.TemporaryDirectory() as tmpdir:
            fpath = os.path.join(tmpdir, "test_admin.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Все КЦР"
            ws.cell(2, 1, "Уникальны обезличинный идентификатор")
            ws.cell(2, 11, "ФИО главного админа")
            ws.cell(2, 12, "Телефон админа")
            ws.cell(2, 13, "Email админа")
            ws.cell(2, 15, "стаус на 21.05.26")
            ws.cell(2, 22, "Всего документов весь 2025")
            ws.cell(2, 23, "Долг весь 2025 на 21.05.26")
            ws.cell(2, 24, "% Долга весь 2025 на 21.05.26")
            ws.cell(2, 25, "Всего документов 2-4 кв 2025")
            ws.cell(2, 26, "Долг 2-4 кв 2025 на 21.05.26")
            ws.cell(2, 27, "% Долга 2-4кв 2025 на 21.05.26")
            ws.cell(2, 28, "Всего документов 1 кв 2026")
            ws.cell(2, 29, "Долг 1 кв 2026 на 21.05.26")
            ws.cell(2, 30, "% Долга 1 кв 2026 на 21.05.26")
            ws.cell(2, 31, "Всего документов 2 кв 2025-1 кв 2026 на 21.05.26")
            ws.cell(2, 32, "Долг 2 кв 2025-1 кв 2026 на 21.05.26")
            ws.cell(2, 33, "% Долга 2 кв 2025-1 кв 2026 на 21.05.26")
            ws.cell(2, 35, "Комментарии ООВС")
            ws.cell(2, 36, "Статус звонка")

            ws.cell(3, 1, "CID1")
            ws.cell(3, 11, "Администратор Иванов")
            ws.cell(3, 12, "+7 (999) 777-88-99")
            ws.cell(3, 13, "admin@cid1.ru")
            wb.save(fpath)
            wb.close()

            HistoricalFilesParser.parse_q1_2026(fpath, known_clients={"CID1"}, clients_dict=clients_dict)
            c = clients_dict["CID1"]
            self.assertEqual(c.manager_name, "Менеджер Сидоров")
            self.assertEqual(c.admin_name, "Администратор Иванов")
            self.assertEqual(c.admin_phone, "+7 (999) 777-88-99")
            self.assertEqual(c.attribute_provenance["admin_name"].source_column, "K")
            self.assertEqual(c.attribute_provenance["admin_phone"].source_column, "L")

    def test_07_statuses_from_base_registry_and_q1_extracted(self):
        """Извлечение статусов колонок F, G основной выгрузки и O базы 2026 года обязательно проверяется утверждениями."""
        with tempfile.TemporaryDirectory() as tmpdir:
            f1_tmp = os.path.join(tmpdir, "base_reg.xlsx")
            wb1 = openpyxl.Workbook()
            ws1 = wb1.active
            ws1.title = "Все КЦР"
            ws1.cell(2, 1, "ID")
            ws1.cell(2, 6, "Статус на 21.07.26")
            ws1.cell(2, 7, "стаус на 21.05.26")
            ws1.cell(2, 10, "Действует ли КЦР")
            ws1.cell(2, 11, "Блокировка в сервисе")
            for c_tot, c_uns, c_pct, _, _, _, p_tok, d_tok in BaseRegistryParser.SLICE_SPECS:
                ws1.cell(2, c_tot, f"Всего документов {p_tok}")
                ws1.cell(2, c_uns, f"Долг {p_tok} на {d_tok}")
                ws1.cell(2, c_pct, f"% Долга {p_tok} на {d_tok}")
            ws1.cell(3, 1, "JAM19209")
            ws1.cell(3, 6, "Долг, пнула МПП")
            ws1.cell(3, 7, "Не нужно работать")
            wb1.save(f1_tmp)
            wb1.close()

            f4_tmp = os.path.join(tmpdir, "q1_2026.xlsx")
            wb4 = openpyxl.Workbook()
            ws4 = wb4.active
            ws4.title = "Все КЦР"
            ws4.cell(2, 1, "Уникальны обезличинный идентификатор")
            ws4.cell(2, 15, "стаус на 21.05.26")
            ws4.cell(2, 22, "Всего документов весь 2025")
            ws4.cell(2, 23, "Долг весь 2025 на 21.05.26")
            ws4.cell(2, 24, "% Долга весь 2025 на 21.05.26")
            ws4.cell(2, 25, "Всего документов 2-4 кв 2025")
            ws4.cell(2, 26, "Долг 2-4 кв 2025 на 21.05.26")
            ws4.cell(2, 27, "% Долга 2-4кв 2025 на 21.05.26")
            ws4.cell(2, 28, "Всего документов 1 кв 2026")
            ws4.cell(2, 29, "Долг 1 кв 2026 на 21.05.26")
            ws4.cell(2, 30, "% Долга 1 кв 2026 на 21.05.26")
            ws4.cell(2, 31, "Всего документов 2 кв 2025-1 кв 2026 на 21.05.26")
            ws4.cell(2, 32, "Долг 2 кв 2025-1 кв 2026 на 21.05.26")
            ws4.cell(2, 33, "% Долга 2 кв 2025-1 кв 2026 на 21.05.26")
            ws4.cell(2, 35, "Комментарии ООВС")
            ws4.cell(2, 36, "Статус звонка")
            ws4.cell(3, 1, "WQC24907")
            ws4.cell(3, 15, "Была претензия")
            wb4.save(f4_tmp)
            wb4.close()

            _, _, h1_syn, _ = BaseRegistryParser.parse(f1_tmp)
            _, h4_syn, _ = HistoricalFilesParser.parse_q1_2026(f4_tmp)
            self.assertEqual(len(h1_syn), 2)
            self.assertEqual(h1_syn[0].occurred_at, "2026-07-21")
            self.assertEqual(h1_syn[0].provenance.source_column, "F")
            self.assertEqual(h1_syn[1].occurred_at, "2026-05-21")
            self.assertEqual(h1_syn[1].provenance.source_column, "G")
            self.assertEqual(len(h4_syn), 1)
            self.assertEqual(h4_syn[0].occurred_at, "2026-05-21")
            self.assertEqual(h4_syn[0].provenance.source_column, "O")

        # Если исходные таблицы присутствуют в окружении — дополнительно сверяем точные количества 7 + 2748 + 115 = 2870
        f1_real = locate_excel_file("КЦР_последняя выгрузка_Дубна.xlsx")
        f4_real = locate_excel_file("База I кв 2026 Дубна.xlsx")
        if f1_real and f4_real:
            _, _, h1, _ = BaseRegistryParser.parse(f1_real)
            _, h4, _ = HistoricalFilesParser.parse_q1_2026(f4_real)
            f_count = sum(1 for h in h1 if h.provenance.source_column == "F" and h.event_type == "historical_status")
            g_count = sum(1 for h in h1 if h.provenance.source_column == "G" and h.event_type == "historical_status")
            o_count = sum(1 for h in h4 if h.provenance.source_column == "O" and h.event_type == "historical_status")
            self.assertEqual(f_count, 7)
            self.assertEqual(g_count, 2748)
            self.assertEqual(o_count, 115)
            self.assertEqual(f_count + g_count + o_count, 2870)

    def test_08_transactional_full_set_recovery(self):
        """При ошибке публикации восстанавливается весь предыдущий согласованный комплект файлов целиком."""
        with tempfile.TemporaryDirectory() as tmpdir:
            export_artifacts([{"id": "OLD"}], [], [], [], [], output_dir=tmpdir)
            real_replace = os.replace
            call_count = {"n": 0}

            def fail_on_eighth_call(src, dst):
                call_count["n"] += 1
                if call_count["n"] == 8:
                    raise OSError("Simulated interruption during staging commit")
                return real_replace(src, dst)

            with patch("os.replace", side_effect=fail_on_eighth_call):
                with self.assertRaises(OSError):
                    export_artifacts([{"id": "NEW"}], [], [], [], [], output_dir=tmpdir)

            with open(os.path.join(tmpdir, "normalized_clients.json"), encoding="utf-8") as f:
                data = json.load(f)
            self.assertEqual(data, [{"id": "OLD"}], "Комплект обязан полностью остаться в согласованном предыдущем состоянии")

    def test_09_formula_and_orphan_percentage_diagnostics(self):
        """Формулы без значений и проценты без количеств диагностируются в отчёте качества."""
        with tempfile.TemporaryDirectory() as tmpdir:
            fpath = os.path.join(tmpdir, "test_diag.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Все КЦР"
            ws.cell(2, 1, "ID")
            ws.cell(2, 6, "Статус на 21.07.26")
            ws.cell(2, 7, "стаус на 21.05.26")
            ws.cell(2, 10, "Действует ли КЦР")
            ws.cell(2, 11, "Блокировка в сервисе")
            for c_tot, c_uns, c_pct, _, _, _, p_tok, d_tok in BaseRegistryParser.SLICE_SPECS:
                ws.cell(2, c_tot, f"Всего документов {p_tok}")
                ws.cell(2, c_uns, f"Долг {p_tok} на {d_tok}")
                ws.cell(2, c_pct, f"% Долга {p_tok} на {d_tok}")

            ws.cell(3, 1, "CID_FORMULA")
            ws.cell(3, 15, 10)
            ws.cell(3, 16, "=2+2")
            ws.cell(4, 1, "CID_ORPHAN")
            ws.cell(4, 17, 0.5)
            wb.save(fpath)
            wb.close()

            _, _, _, issues = BaseRegistryParser.parse(fpath)
            codes = {i.code for i in issues}
            self.assertIn("INVALID_NUMERIC_UNSENT", codes)
            self.assertIn("ORPHAN_PERCENTAGE_WITHOUT_COUNTS", codes)

    def test_10_date_and_period_bounds_validation(self):
        """Проверка календарной корректности дат и условия period_start <= period_end."""
        bad_date = SliceRecord("C1", "2025-01-01", "2025-12-31", "2026-99-99", 10, 1, None, Provenance("f", "s", 1))
        bad_bounds = SliceRecord("C2", "2026-06-01", "2026-01-01", "2026-10-20", 10, 1, None, Provenance("f", "s", 2))
        valid, rejected, issues = QualityAuditor.audit_and_filter_slices([bad_date, bad_bounds])
        self.assertEqual(len(valid), 0)
        self.assertEqual(len(rejected), 2)
        codes = {i.code for i in issues}
        self.assertIn("INVALID_DATE_VALUE", codes)
        self.assertIn("INVALID_PERIOD_BOUNDS", codes)

    def test_11_preflight_and_unknown_client_isolation(self):
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

    def test_12_verbatim_history_preserved(self):
        """Телефоны, email и ФИО сохраняются дословно без маскирования как при разборе Excel, так и в normalized_history.jsonl."""
        with tempfile.TemporaryDirectory() as tmpdir:
            f3_tmp = os.path.join(tmpdir, "debts_2025.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "1-3 кв 25"
            ws.cell(1, 1, "Уникальны обезличинный идентификатор")
            ws.cell(1, 10, "коммаентарий для претензии")
            for col_idx, date_tok, _ in HistoricalFilesParser.STATUS_COLUMNS_SPEC:
                ws.cell(1, col_idx, f"Статус на {date_tok}")
            ws.cell(1, 19, "Документов всего за 2025 год")
            ws.cell(1, 20, "Долг на 27.03.26")
            ws.cell(1, 21, "% Долга на 27.03.26")
            ws.cell(1, 23, "долг на 16.01.25")
            ws.cell(1, 24, "% Долга на 16.01.26")
            ws.cell(1, 27, "долг на 05.12.25")
            ws.cell(1, 28, "% Долга на 05.12.25")
            ws.cell(1, 31, "долг на 05.11.25")
            ws.cell(1, 32, "% Долга на 05.11.25")
            ws.cell(1, 33, "долг на 31.10.25")
            ws.cell(1, 34, "% Долга на 31.10.25")
            ws.cell(1, 35, "долг на 20.10.25")
            ws.cell(1, 46, "Комментарии ООВС")
            ws.cell(1, 47, "Статус звонка")
            ws.cell(1, 53, "Комментарий к повторному звонку")
            ws.cell(1, 54, "Статус звонка в январе")

            ws.cell(2, 1, "CID_VERB")
            ws.cell(2, 46, "it - 79995209332 Вантеев, +7 (914) 200-06-94, admin@kontur.ru")
            ws.cell(2, 47, "Прошел первый звонок")
            wb.save(f3_tmp)
            wb.close()

            _, h_syn, _ = HistoricalFilesParser.parse_debts_2025(f3_tmp)
            self.assertEqual(len(h_syn), 1)
            self.assertIn("79995209332", h_syn[0].text)
            self.assertIn("+7 (914) 200-06-94", h_syn[0].text)
            self.assertIn("admin@kontur.ru", h_syn[0].text)

        hist_jsonl = os.path.join(SCRIPT_DIR, "normalized_history.jsonl")
        self.assertTrue(os.path.exists(hist_jsonl), "Файл normalized_history.jsonl должен существовать")
        with open(hist_jsonl, encoding="utf-8") as f:
            content = f.read()
        self.assertIn("79995209332", content)
        self.assertIn("+7 (914) 200-06-94", content)
        self.assertNotIn("[ТЕЛЕФОН_СКРЫТ]", content)

    def test_13_repeated_slice_without_total_allowed_when_period_total_known(self):
        """Повторное наблюдение без total на новую дату допускается, если для (client, period) уже известен валидный total."""
        primary = SliceRecord(
            "C_REP", "2025-01-01", "2025-12-31", "2026-05-21",
            total=10, unsent=5, raw_percentage=0.5,
            provenance=Provenance("f1.xlsx", "Все КЦР", 10, "O,V")
        )
        repeated_ok = SliceRecord(
            "C_REP", "2025-01-01", "2025-12-31", "2026-10-20",
            total=None, unsent=2, raw_percentage=None,
            provenance=Provenance("f2.xlsx", "весь 2025 год", 15, "D,E")
        )
        repeated_exceeds = SliceRecord(
            "C_REP", "2025-01-01", "2025-12-31", "2026-11-20",
            total=None, unsent=15, raw_percentage=None,
            provenance=Provenance("f2.xlsx", "весь 2025 год", 16, "D,E")
        )
        orphan_no_total = SliceRecord(
            "C_NO_TOT", "2025-04-01", "2026-03-31", "2026-05-21",
            total=None, unsent=0, raw_percentage=None,
            provenance=Provenance("f1.xlsx", "Все КЦР", 20, "AN,AS")
        )
        valid, rejected, _ = QualityAuditor.audit_and_filter_slices([primary, repeated_ok, repeated_exceeds, orphan_no_total])
        self.assertEqual(len(valid), 2)
        self.assertIsNone(valid[1].total)
        self.assertEqual(valid[1].unsent, 2)
        self.assertEqual(len(rejected), 2)
        rej_reasons = {r.source_client_id: r.rejection_reason for r in rejected}
        self.assertEqual(rej_reasons["C_REP"], "UNSENT_EXCEEDS_TOTAL")
        self.assertEqual(rej_reasons["C_NO_TOT"], "UNSENT_WITHOUT_TOTAL")

    def test_14_column_mapping_completeness_all_188_columns(self):
        """Полная карта колонок покрывает все 188 столбцов 7 листов без пропусков и с обоснованием исключений."""
        expected_counts = {
            "КЦР_последняя выгрузка_Дубна.xlsx::Все КЦР": 66,
            "Срез 20.10.26_Дубна.xlsx::весь 2025 год": 5,
            "Срез 20.10.26_Дубна.xlsx::1 кв 2026": 5,
            "Срез 20.10.26_Дубна.xlsx::2 кв 2026": 5,
            "Срез 20.10.26_Дубна.xlsx::3 кв 2026": 5,
            "База должники III кв 2025 Дубна.xlsx::1-3 кв 25": 60,
            "База I кв 2026 Дубна.xlsx::Все КЦР": 42,
        }
        total_cols = 0
        for sheet_key, exp_len in expected_counts.items():
            self.assertIn(sheet_key, COLUMN_MAPPING)
            cols = COLUMN_MAPPING[sheet_key]
            self.assertEqual(len(cols), exp_len)
            indices = [c["col_index"] for c in cols]
            self.assertEqual(indices, list(range(1, exp_len + 1)))
            for c in cols:
                if c["target_field"] is None:
                    self.assertTrue(
                        c["exclusion_reason"] and len(c["exclusion_reason"]) > 5,
                        f"Для исключённой колонки {sheet_key}::{c['col_letter']} обязательна причина исключения"
                    )
            total_cols += len(cols)
        self.assertEqual(total_cols, 188)

    @staticmethod
    def _reference_slice(as_of, total, unsent, percentage=None, reason=None):
        return SliceRecord(
            "REFERENCE_TEST", "2025-01-01", "2025-12-31", as_of,
            total, unsent, percentage, Provenance("synthetic.xlsx", "test", 2, "D,E"), reason
        )

    def test_15_conflicting_references_cannot_support_missing_total(self):
        records = [
            self._reference_slice("2026-05-21", 10, 2),
            self._reference_slice("2026-05-21", 10, 3),
            self._reference_slice("2026-10-20", None, 1),
        ]
        for order in itertools.permutations(records):
            with self.subTest(order=[(s.as_of_date, s.unsent) for s in order]):
                valid, rejected, _ = QualityAuditor.audit_and_filter_slices(copy.deepcopy(order))
                self.assertEqual(valid, [])
                self.assertEqual(len(rejected), 3)
                repeated = next(s for s in rejected if s.total is None)
                self.assertEqual(repeated.rejection_reason, "UNSENT_WITHOUT_TOTAL")
                self.assertEqual(sum(s.rejection_reason == "CONFLICTING_DUPLICATE_SLICE" for s in rejected), 2)

    def test_16_rejected_reference_does_not_override_valid_alternative(self):
        records = [
            self._reference_slice("2026-05-21", 20, 2),
            self._reference_slice("2026-05-21", 20, 3),
            self._reference_slice("2026-07-21", 10, 4),
            self._reference_slice("2026-10-20", None, 1),
        ]
        for order in itertools.permutations(records):
            valid, rejected, _ = QualityAuditor.audit_and_filter_slices(copy.deepcopy(order))
            self.assertEqual({s.as_of_date for s in valid}, {"2026-07-21", "2026-10-20"})
            self.assertEqual(len(rejected), 2)
            self.assertIsNone(next(s for s in valid if s.as_of_date == "2026-10-20").total)

    def test_17_invalid_and_blocked_references_cannot_support_missing_total(self):
        for percentage, reason in [(float("nan"), None), (None, "UNKNOWN_CLIENT_ID")]:
            records = [
                self._reference_slice("2026-05-21", 10, 2, percentage, reason),
                self._reference_slice("2026-10-20", None, 1),
            ]
            for order in itertools.permutations(records):
                valid, rejected, _ = QualityAuditor.audit_and_filter_slices(copy.deepcopy(order))
                self.assertEqual(valid, [])
                self.assertEqual(len(rejected), 2)
                self.assertEqual(next(s for s in rejected if s.total is None).rejection_reason, "UNSENT_WITHOUT_TOTAL")

    def test_18_removed_reference_is_rechecked_after_final_conflicts(self):
        # Повторы конфликтуют с обеими предварительными опорами.
        # Запись третьей даты лишается последней принятой опоры.
        records = [
            self._reference_slice("2026-05-21", 10, 2),
            self._reference_slice("2026-07-21", 10, 3),
            self._reference_slice("2026-05-21", None, 1),
            self._reference_slice("2026-07-21", None, 1),
            self._reference_slice("2026-10-20", None, 1),
        ]
        for order in itertools.permutations(records):
            valid, rejected, _ = QualityAuditor.audit_and_filter_slices(copy.deepcopy(order))
            self.assertEqual(valid, [])
            self.assertEqual(len(rejected), 5)
            self.assertEqual(next(s for s in rejected if s.as_of_date == "2026-10-20").rejection_reason,
                             "UNSENT_WITHOUT_TOTAL")

    def test_19_identical_duplicates_keep_reference_but_totals_must_be_unambiguous(self):
        records = [
            self._reference_slice("2026-05-21", 10, 2),
            self._reference_slice("2026-05-21", 10, 2),
            self._reference_slice("2026-10-20", None, 1),
        ]
        valid, rejected, _ = QualityAuditor.audit_and_filter_slices(copy.deepcopy(records))
        self.assertEqual(len(valid), 2)
        self.assertEqual([s.rejection_reason for s in rejected], ["DUPLICATE_SLICE"])
        records.append(self._reference_slice("2026-07-21", 20, 3))
        valid, rejected, _ = QualityAuditor.audit_and_filter_slices(copy.deepcopy(records))
        self.assertEqual(len(valid), 2)
        self.assertEqual(next(s for s in rejected if s.total is None).rejection_reason, "UNSENT_WITHOUT_TOTAL")

    def test_20_creation_date_is_calendar_date_or_null(self):
        for raw, expected in [("начало 2026 года", None), ("31.02.2026", None),
                              ("31.01.2026", "2026-01-31"), (None, None)]:
            with self.subTest(raw=raw), tempfile.TemporaryDirectory() as temp:
                path = os.path.join(temp, "creation_date.xlsx")
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "1-3 кв 25"
                headers = {1: "идентификатор", 10: "претензии", 19: "документов всего за 2025 год",
                           20: "долг 27.03.26", 27: "долг 05.12.25", 31: "долг 05.11.25",
                           33: "долг 31.10.25", 35: "долг 20.10.25", 46: "комментарии оовс",
                           47: "статус звонка", 53: "повторному звонку", 54: "статус звонка в январе"}
                for col, dt, _ in HistoricalFilesParser.STATUS_COLUMNS_SPEC:
                    headers[col] = "статус " + dt
                for col, value in headers.items():
                    ws.cell(1, col, value)
                ws.cell(2, 1, "DATE_TEST")
                ws.cell(2, 4, raw)
                wb.save(path)
                wb.close()
                client = ClientRecord("DATE_TEST", None, None, None, None, Provenance("f", "s", 1))
                _, _, issues = HistoricalFilesParser.parse_debts_2025(
                    path, {"DATE_TEST"}, {"DATE_TEST": client})
                self.assertEqual(client.creation_date, expected)
                self.assertEqual(client.raw_creation_date, raw)
                errors = [q for q in issues if q.code == "INVALID_CREATION_DATE"]
                self.assertEqual(len(errors), int(raw is not None and expected is None))
                if raw is not None:
                    self.assertEqual(client.attribute_provenance["raw_creation_date"].source_column, "D")
                if errors:
                    self.assertEqual((errors[0].row, errors[0].col, errors[0].raw_value), (2, "D", raw))

    def test_21_service_center_code_and_name_are_both_preserved(self):
        for code, name in [("SC-42", "СЦ Восток"), ("SC-42", None), (None, "СЦ Восток")]:
            with self.subTest(code=code, name=name), tempfile.TemporaryDirectory() as temp:
                path = os.path.join(temp, "service_center.xlsx")
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "1 кв 2026"
                for col, value in enumerate(["идентификатор", "код СЦ", "название СЦ",
                                              "документов всего", "не отправлено"], 1):
                    ws.cell(1, col, value)
                ws.append(["SC_TEST", code, name, 10, 2])
                wb.save(path)
                wb.close()
                client = ClientRecord("SC_TEST", None, None, None, None, Provenance("f", "s", 1))
                PeriodicSliceParser.parse(path, {"SC_TEST"}, clients_dict={"SC_TEST": client})
                self.assertEqual(client.service_center_code, code)
                self.assertEqual(client.service_center, name)
                self.assertIsNone(client.kcr_code)
                if code:
                    self.assertEqual(client.attribute_provenance["service_center_code"].source_column, "B")
                if name:
                    self.assertEqual(client.attribute_provenance["service_center"].source_column, "C")

    def test_22_repeated_percentage_is_checked_against_accepted_reference(self):
        primary = self._reference_slice("2026-05-21", 10, 5, 0.5)
        repeated = self._reference_slice("2026-10-20", None, 2, 0.9)
        for order in itertools.permutations([primary, repeated]):
            valid, rejected, issues = QualityAuditor.audit_and_filter_slices(copy.deepcopy(order))
            self.assertEqual(len(valid), 2)
            self.assertEqual(rejected, [])
            self.assertIsNone(next(s for s in valid if s.as_of_date == "2026-10-20").total)
            mismatches = [q for q in issues if q.code == "PERCENTAGE_MISMATCH"]
            self.assertEqual(len(mismatches), 1)
            self.assertEqual(mismatches[0].raw_value, "excel=0.9000, calc=0.2000")

    def test_23_repeated_percentage_correct_missing_and_zero_total(self):
        for total, unsent, percentage in [(10, 2, 0.2), (10, 2, None), (0, 0, 0.9)]:
            records = [self._reference_slice("2026-05-21", total, 0),
                       self._reference_slice("2026-10-20", None, unsent, percentage)]
            valid, rejected, issues = QualityAuditor.audit_and_filter_slices(records)
            self.assertEqual(len(valid), 2)
            self.assertEqual(rejected, [])
            self.assertFalse(any(q.code == "PERCENTAGE_MISMATCH" for q in issues))

    def test_24_rejected_references_do_not_trigger_percentage_comparison(self):
        records = [self._reference_slice("2026-05-21", 10, 2),
                   self._reference_slice("2026-05-21", 10, 3),
                   self._reference_slice("2026-10-20", None, 2, 0.9)]
        valid, rejected, issues = QualityAuditor.audit_and_filter_slices(records)
        self.assertEqual(valid, [])
        self.assertEqual(len(rejected), 3)
        self.assertFalse(any(q.code == "PERCENTAGE_MISMATCH" for q in issues))

    def test_25_repeated_percentage_uses_final_surviving_reference(self):
        records = [self._reference_slice("2026-05-21", 10, 5),
                   self._reference_slice("2026-07-21", 10, 4),
                   self._reference_slice("2026-05-21", None, 1),
                   self._reference_slice("2026-10-20", None, 2, 0.9)]
        for order in itertools.permutations(records):
            valid, _, issues = QualityAuditor.audit_and_filter_slices(copy.deepcopy(order))
            self.assertEqual({s.as_of_date for s in valid}, {"2026-07-21", "2026-10-20"})
            self.assertEqual(sum(q.code == "PERCENTAGE_MISMATCH" for q in issues), 1)


if __name__ == "__main__":
    unittest.main()
