import os
import json
import unittest

DIR = os.path.dirname(os.path.abspath(__file__))

class TestKCRParser(unittest.TestCase):
    def test_01_files_exist(self):
        """Проверка наличия всех требуемых выходных файлов."""
        expected = [
            "normalized_clients.json",
            "normalized_slices.jsonl",
            "normalized_history.jsonl",
            "quality_report.json",
            "quality_report.csv"
        ]
        for fname in expected:
            path = os.path.join(DIR, fname)
            self.assertTrue(os.path.exists(path), f"Файл {fname} отсутствует!")

    def test_02_clients_contract(self):
        """Проверка контракта клиентов (5874 записи, источник, поля)."""
        path = os.path.join(DIR, "normalized_clients.json")
        with open(path, "r", encoding="utf-8") as f:
            clients = json.load(f)
        self.assertEqual(len(clients), 5874)
        for c in clients[:100]:
            self.assertIn("source_client_id", c)
            self.assertIn("provenance", c)
            self.assertEqual(c["provenance"]["source_file"], "КЦР_последняя выгрузка_Дубна.xlsx")
            self.assertGreaterEqual(c["provenance"]["source_row"], 3)

    def test_03_slices_dates_and_nulls(self):
        """Проверка формата дат (YYYY-MM-DD) и сохранения честных null (не заменены нулями)."""
        path = os.path.join(DIR, "normalized_slices.jsonl")
        has_null_unsent = False
        count = 0
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                count += 1
                item = json.loads(line)
                self.assertIn("source_client_id", item)
                self.assertEqual(len(item["period_start"]), 10)
                self.assertEqual(len(item["period_end"]), 10)
                self.assertEqual(len(item["as_of_date"]), 10)
                if item["unsent"] is None:
                    has_null_unsent = True
        self.assertEqual(count, 45204)
        self.assertTrue(has_null_unsent, "Пустые значения unsent должны оставаться null, а не превращаться в 0")

    def test_04_history_materials(self):
        """Проверка истории коммуникаций."""
        path = os.path.join(DIR, "normalized_history.jsonl")
        count = 0
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                count += 1
                item = json.loads(line)
                self.assertIn("source_client_id", item)
                self.assertIn("event_type", item)
                self.assertIn("provenance", item)
        self.assertEqual(count, 410)

    def test_05_quality_report(self):
        """Проверка выявления дублей, неизвестных клиентов и служебных строк."""
        path = os.path.join(DIR, "quality_report.json")
        with open(path, "r", encoding="utf-8") as f:
            errors = json.load(f)
        self.assertEqual(len(errors), 249)
        codes = {e["code"] for e in errors}
        self.assertIn("DUPLICATE_CLIENT_ID", codes)
        self.assertIn("UNKNOWN_CLIENT_ID", codes)
        self.assertIn("MISSING_CLIENT_ID", codes)

if __name__ == "__main__":
    unittest.main()
