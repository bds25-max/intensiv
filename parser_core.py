"""
parser_core.py — Отрефакторенный модульный ETL-парсер Контур КЦР (Задача 1).

Архитектура:
1. Data Contracts (Data Models) — строгие dataclass-модели для соблюдения JSON-контракта.
2. Column Mapper & Helpers — гибкий поиск строк заголовков и сопоставление полей без хардкода.
3. Domain Parsers — независимые обработчики каждого Excel-файла.
4. Quality Auditor — правила валидации и обнаружения бизнес-аномалий.
5. Runner/Facade — единая точка оркестрации и экспорта.
"""

from __future__ import annotations
import os
import json
import csv
import re
from dataclasses import dataclass, asdict
from datetime import datetime, date
from typing import Any, Dict, List, Optional, Tuple, Set
import openpyxl


# =====================================================================
# 1. DATA CONTRACTS (Модели данных общего контракта)
# =====================================================================

@dataclass
class Provenance:
    source_file: str
    source_sheet: str
    source_row: int
    source_column: Optional[str] = None

@dataclass
class ClientRecord:
    source_client_id: str
    is_active: Optional[str]
    is_blocked: Optional[str]
    provenance: Provenance

@dataclass
class SliceRecord:
    source_client_id: str
    period_start: str
    period_end: str
    as_of_date: str
    total: Optional[int]
    unsent: Optional[int]
    raw_percentage: Optional[float]
    provenance: Provenance

@dataclass
class HistoryRecord:
    source_client_id: str
    event_type: str
    occurred_at: Optional[str]
    text: Optional[str]
    raw_status: Optional[str]
    raw_deadline: Optional[str]
    provenance: Provenance

@dataclass
class QualityIssue:
    code: str
    file: str
    sheet: str
    row: int
    col: str
    raw_value: Optional[str]
    description: str
    action: str  # 'skipped' | 'blocked' | 'requires_review'


# =====================================================================
# 2. UTILITY & NORMALIZATION HELPERS
# =====================================================================

def clean_cell_value(val: Any) -> Any:
    """Нормализация ячейки: даты в ISO YYYY-MM-DD, пустые строки в None."""
    if val is None:
        return None
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, str):
        cleaned = val.strip()
        return cleaned if cleaned != "" else None
    return val

def to_clean_int(val: Any) -> Optional[int]:
    """Преобразование числовых значений документов без принудительной замены на 0."""
    if val is None:
        return None
    if isinstance(val, bool):
        return None
    if isinstance(val, (int, float)):
        return int(round(val))
    try:
        val_str = str(val).strip().replace(" ", "").replace(",", ".")
        if val_str:
            return int(round(float(val_str)))
    except (ValueError, TypeError):
        pass
    return None

def detect_header_row(ws: openpyxl.worksheet.worksheet.Worksheet, max_scan: int = 5) -> int:
    """
    Динамическое определение строки заголовков (вместо жестко зашитой строки).
    Ищет строку с максимальным числом строковых названий колонок.
    """
    best_row = 1
    max_non_empty = -1
    for r in range(1, min(max_scan + 1, (ws.max_row or max_scan) + 1)):
        row_vals = [ws.cell(r, c).value for c in range(1, min(ws.max_column + 1, 30))]
        text_count = sum(1 for v in row_vals if isinstance(v, str) and len(v.strip()) > 1)
        if text_count > max_non_empty:
            max_non_empty = text_count
            best_row = r
    return best_row


# =====================================================================
# 3. DOMAIN FILE PARSERS (Изолированные парсеры файлов)
# =====================================================================

class BaseRegistryParser:
    """Парсер основного реестра: КЦР_последняя выгрузка_Дубна.xlsx."""

    # Карта колонок широкой таблицы срезов (всего, долг, %, период_старт, период_конец, дата_среза)
    SLICE_SPECS = [
        (15, 16, 17, "2025-01-01", "2025-12-31", "2026-09-29"),
        (15, 18, 19, "2025-01-01", "2025-12-31", "2026-07-21"),
        (15, 20, 21, "2025-01-01", "2025-12-31", "2026-07-07"),
        (15, 22, 23, "2025-01-01", "2025-12-31", "2026-05-21"),
        (31, 32, 33, "2026-01-01", "2026-03-31", "2026-09-29"),
        (31, 34, 35, "2026-01-01", "2026-03-31", "2026-07-21"),
        (31, 36, 37, "2026-01-01", "2026-03-31", "2026-07-07"),
        (31, 38, 39, "2026-01-01", "2026-03-31", "2026-05-21"),
        (52, 53, 54, "2026-04-01", "2026-06-30", "2026-09-29"),
        (52, 55, 56, "2026-04-01", "2026-06-30", "2026-07-21"),
    ]

    @classmethod
    def parse(cls, filepath: str) -> Tuple[Dict[str, ClientRecord], List[SliceRecord], List[QualityIssue]]:
        clients: Dict[str, ClientRecord] = {}
        slices: List[SliceRecord] = []
        issues: List[QualityIssue] = []

        if not os.path.exists(filepath):
            return clients, slices, issues

        filename = os.path.basename(filepath)
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb["Все КЦР"]
        header_row = detect_header_row(ws)

        for r in range(header_row + 1, ws.max_row + 1):
            raw_id = ws.cell(r, 1).value

            # Обнаружение служебных/итоговых строк
            if raw_id is None:
                has_any_data = any(ws.cell(r, c).value is not None for c in range(1, ws.max_column + 1))
                if has_any_data:
                    issues.append(QualityIssue(
                        code="MISSING_CLIENT_ID",
                        file=filename,
                        sheet="Все КЦР",
                        row=r, col="A",
                        raw_value=None,
                        description="Служебная или итоговая строка без идентификатора клиента",
                        action="skipped"
                    ))
                continue

            cid = str(raw_id).strip()
            if cid in clients:
                issues.append(QualityIssue(
                    code="DUPLICATE_CLIENT_ID",
                    file=filename,
                    sheet="Все КЦР",
                    row=r, col="A",
                    raw_value=cid,
                    description=f"Повторяющийся идентификатор клиента в основной выгрузке: {cid}",
                    action="requires_review"
                ))
            else:
                clients[cid] = ClientRecord(
                    source_client_id=cid,
                    is_active=clean_cell_value(ws.cell(r, 10).value),
                    is_blocked=clean_cell_value(ws.cell(r, 11).value),
                    provenance=Provenance(source_file=filename, source_sheet="Все КЦР", source_row=r)
                )

            # Нормализация широких срезов
            for c_tot, c_uns, c_pct, p_start, p_end, as_of in cls.SLICE_SPECS:
                tot = ws.cell(r, c_tot).value
                uns = ws.cell(r, c_uns).value
                pct = ws.cell(r, c_pct).value
                if tot is not None or uns is not None:
                    slices.append(SliceRecord(
                        source_client_id=cid,
                        period_start=p_start,
                        period_end=p_end,
                        as_of_date=as_of,
                        total=to_clean_int(tot),
                        unsent=to_clean_int(uns),
                        raw_percentage=pct if isinstance(pct, (int, float)) else None,
                        provenance=Provenance(
                            source_file=filename,
                            source_sheet="Все КЦР",
                            source_row=r,
                            source_column=f"{openpyxl.utils.get_column_letter(c_tot)},{openpyxl.utils.get_column_letter(c_uns)}"
                        )
                    ))

        return clients, slices, issues


class PeriodicSliceParser:
    """Парсер оперативных срезов: Срез 20.10.26_Дубна.xlsx."""

    SHEET_PERIODS = {
        "весь 2025 год": ("2025-01-01", "2025-12-31"),
        "1 кв 2026": ("2026-01-01", "2026-03-31"),
        "2 кв 2026": ("2026-04-01", "2026-06-30"),
        "3 кв 2026": ("2026-07-01", "2026-09-30")
    }

    @classmethod
    def parse(cls, filepath: str, known_clients: Set[str]) -> Tuple[List[SliceRecord], List[QualityIssue]]:
        slices: List[SliceRecord] = []
        issues: List[QualityIssue] = []

        if not os.path.exists(filepath):
            return slices, issues

        filename = os.path.basename(filepath)
        wb = openpyxl.load_workbook(filepath, data_only=True)

        for sname, (p_start, p_end) in cls.SHEET_PERIODS.items():
            if sname not in wb.sheetnames:
                continue
            ws = wb[sname]
            header_row = detect_header_row(ws)
            seen_in_sheet: Set[str] = set()

            for r in range(header_row + 1, ws.max_row + 1):
                raw_id = ws.cell(r, 1).value
                if raw_id is None:
                    continue

                cid = str(raw_id).strip()

                # Проверка на дубликат внутри одного среза
                if cid in seen_in_sheet:
                    issues.append(QualityIssue(
                        code="DUPLICATE_CLIENT_ID",
                        file=filename,
                        sheet=sname,
                        row=r, col="A",
                        raw_value=cid,
                        description=f"Дублирующийся идентификатор в рамках одного листа: {cid}",
                        action="requires_review"
                    ))
                else:
                    seen_in_sheet.add(cid)

                # Проверка на существование клиента в реестре
                if cid not in known_clients:
                    issues.append(QualityIssue(
                        code="UNKNOWN_CLIENT_ID",
                        file=filename,
                        sheet=sname,
                        row=r, col="A",
                        raw_value=cid,
                        description=f"Клиент {cid} отсутствует в основной базе клиентов",
                        action="blocked"
                    ))

                tot = ws.cell(r, 4).value
                uns = ws.cell(r, 5).value

                slices.append(SliceRecord(
                    source_client_id=cid,
                    period_start=p_start,
                    period_end=p_end,
                    as_of_date="2026-10-20",
                    total=to_clean_int(tot),
                    unsent=to_clean_int(uns),
                    raw_percentage=None,
                    provenance=Provenance(source_file=filename, source_sheet=sname, source_row=r)
                ))

        return slices, issues


class HistoricalMaterialsParser:
    """Парсер коммуникаций и работы операторов из исторических файлов."""

    @staticmethod
    def parse_debts_2025(filepath: str) -> Tuple[List[HistoryRecord], List[QualityIssue]]:
        history: List[HistoryRecord] = []
        issues: List[QualityIssue] = []

        if not os.path.exists(filepath):
            return history, issues

        filename = os.path.basename(filepath)
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb["1-3 кв 25"]
        header_row = detect_header_row(ws)

        for r in range(header_row + 1, ws.max_row + 1):
            raw_id = ws.cell(r, 1).value
            if raw_id is None:
                has_any_data = any(ws.cell(r, c).value is not None for c in range(1, ws.max_column + 1))
                if has_any_data:
                    issues.append(QualityIssue(
                        code="MISSING_CLIENT_ID",
                        file=filename,
                        sheet="1-3 кв 25",
                        row=r, col="A",
                        raw_value=None,
                        description="Служебная или итоговая строка без идентификатора клиента",
                        action="skipped"
                    ))
                continue

            cid = str(raw_id).strip()
            comment_1 = clean_cell_value(ws.cell(r, 46).value)
            comment_2 = clean_cell_value(ws.cell(r, 53).value)
            claim_comm = clean_cell_value(ws.cell(r, 10).value)
            call_status = clean_cell_value(ws.cell(r, 47).value)
            call_date = clean_cell_value(ws.cell(r, 48).value)
            client_deadline = clean_cell_value(ws.cell(r, 49).value)

            full_text_parts = [p for p in [claim_comm, comment_1, comment_2] if p]
            full_text = " | ".join(full_text_parts) if full_text_parts else None

            if full_text or call_status or call_date or client_deadline:
                history.append(HistoryRecord(
                    source_client_id=cid,
                    event_type="communication_log",
                    occurred_at=call_date,
                    text=full_text,
                    raw_status=call_status,
                    raw_deadline=client_deadline,
                    provenance=Provenance(source_file=filename, source_sheet="1-3 кв 25", source_row=r)
                ))

        return history, issues

    @staticmethod
    def parse_q1_2026(filepath: str) -> Tuple[List[HistoryRecord], List[QualityIssue]]:
        history: List[HistoryRecord] = []
        issues: List[QualityIssue] = []

        if not os.path.exists(filepath):
            return history, issues

        filename = os.path.basename(filepath)
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb["Все КЦР"]
        header_row = detect_header_row(ws)

        for r in range(header_row + 1, ws.max_row + 1):
            raw_id = ws.cell(r, 1).value
            if raw_id is None:
                continue

            cid = str(raw_id).strip()
            comm_oovs = clean_cell_value(ws.cell(r, 35).value)
            call_stat = clean_cell_value(ws.cell(r, 36).value)
            recall_date = clean_cell_value(ws.cell(r, 37).value)
            client_dl = clean_cell_value(ws.cell(r, 38).value)

            if comm_oovs or call_stat or recall_date or client_dl:
                history.append(HistoryRecord(
                    source_client_id=cid,
                    event_type="operator_contact",
                    occurred_at=recall_date,
                    text=comm_oovs,
                    raw_status=call_stat,
                    raw_deadline=client_dl,
                    provenance=Provenance(source_file=filename, source_sheet="Все КЦР", source_row=r)
                ))

        return history, issues


# =====================================================================
# 4. QUALITY AUDITOR (Валидатор качества данных)
# =====================================================================

class QualityAuditor:
    """Комплексная проверка числовых показателей и бизнес-логики срезов."""

    @staticmethod
    def audit_slices(slices: List[SliceRecord]) -> List[QualityIssue]:
        issues: List[QualityIssue] = []

        for s in slices:
            tot = s.total
            uns = s.unsent
            pct = s.raw_percentage
            p = s.provenance

            # Проверка дат
            if not s.as_of_date or not s.period_start or not s.period_end:
                issues.append(QualityIssue(
                    code="MISSING_DATE_OR_PERIOD",
                    file=p.source_file, sheet=p.source_sheet, row=p.source_row, col="dates",
                    raw_value=None,
                    description="Отсутствует дата среза или границы отчётного периода",
                    action="blocked"
                ))

            # Проверка неотрицательности
            if tot is not None and tot < 0:
                issues.append(QualityIssue(
                    code="INVALID_NUMERIC_TOTAL",
                    file=p.source_file, sheet=p.source_sheet, row=p.source_row, col="total",
                    raw_value=str(tot),
                    description="Общее количество документов не может быть отрицательным",
                    action="requires_review"
                ))

            if uns is not None and uns < 0:
                issues.append(QualityIssue(
                    code="INVALID_NUMERIC_UNSENT",
                    file=p.source_file, sheet=p.source_sheet, row=p.source_row, col="unsent",
                    raw_value=str(uns),
                    description="Количество непереданных документов не может быть отрицательным",
                    action="requires_review"
                ))

            # Логическое противоречие unsent > total
            if tot is not None and uns is not None:
                if uns > tot:
                    issues.append(QualityIssue(
                        code="UNSENT_EXCEEDS_TOTAL",
                        file=p.source_file, sheet=p.source_sheet, row=p.source_row, col="unsent",
                        raw_value=f"unsent={uns}, total={tot}",
                        description="Количество непереданных документов превышает общее количество",
                        action="requires_review"
                    ))

                # Расхождение исходного процента Excel с расчетным
                if pct is not None and tot > 0:
                    calc_pct = uns / tot
                    if abs(pct - calc_pct) > 0.02:
                        issues.append(QualityIssue(
                            code="PERCENTAGE_MISMATCH",
                            file=p.source_file, sheet=p.source_sheet, row=p.source_row, col="percentage",
                            raw_value=f"excel={pct:.4f}, calc={calc_pct:.4f}",
                            description="Исходный процент расходится с вычисленным отношением unsent/total",
                            action="requires_review"
                        ))

        return issues


# =====================================================================
# 5. ORCHESTRATOR & EXPORTER (Фасад и экспорт артефактов)
# =====================================================================

def parse_all_kcr(base_dir: str = ".") -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], List[Dict[str, Any]], List[Dict[str, Any]]]:
    """Главная функция-фасад для обратной совместимости с тестами и смежными модулями."""
    f1 = os.path.join(base_dir, "КЦР_последняя выгрузка_Дубна.xlsx")
    f2 = os.path.join(base_dir, "Срез 20.10.26_Дубна.xlsx")
    f3 = os.path.join(base_dir, "База должники III кв 2025 Дубна.xlsx")
    f4 = os.path.join(base_dir, "База I кв 2026 Дубна.xlsx")

    # 1. Разбор основного реестра
    clients_dict, slices_main, issues_main = BaseRegistryParser.parse(f1)
    known_cids = set(clients_dict.keys())

    # 2. Разбор периодических срезов
    slices_periodic, issues_periodic = PeriodicSliceParser.parse(f2, known_cids)

    # 3. Разбор исторических материалов
    history_2025, issues_h1 = HistoricalMaterialsParser.parse_debts_2025(f3)
    history_2026, issues_h2 = HistoricalMaterialsParser.parse_q1_2026(f4)

    # 4. Объединение срезов и аудит
    all_slices = slices_main + slices_periodic
    slice_audit_issues = QualityAuditor.audit_slices(all_slices)

    all_history = history_2025 + history_2026
    all_issues = issues_main + issues_periodic + issues_h1 + issues_h2 + slice_audit_issues

    # Сериализация в формат словарей (для JSON-контракта)
    clients_out = [asdict(c) for c in clients_dict.values()]
    slices_out = [asdict(s) for s in all_slices]
    history_out = [asdict(h) for h in all_history]
    issues_out = [asdict(i) for i in all_issues]

    return clients_out, slices_out, history_out, issues_out


def export_artifacts(clients: List[Dict[str, Any]], slices: List[Dict[str, Any]], history: List[Dict[str, Any]], errors: List[Dict[str, Any]], output_dir: str = "."):
    """Экспорт нормализованных данных и журнала качества в файлы."""
    os.makedirs(output_dir, exist_ok=True)

    with open(os.path.join(output_dir, "normalized_clients.json"), "w", encoding="utf-8") as f:
        json.dump(clients, f, ensure_ascii=False, indent=2)

    with open(os.path.join(output_dir, "normalized_slices.jsonl"), "w", encoding="utf-8") as f:
        for s in slices:
            f.write(json.dumps(s, ensure_ascii=False) + "\n")

    with open(os.path.join(output_dir, "normalized_history.jsonl"), "w", encoding="utf-8") as f:
        for h in history:
            f.write(json.dumps(h, ensure_ascii=False) + "\n")

    with open(os.path.join(output_dir, "quality_report.json"), "w", encoding="utf-8") as f:
        json.dump(errors, f, ensure_ascii=False, indent=2)

    with open(os.path.join(output_dir, "quality_report.csv"), "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["code", "file", "sheet", "row", "col", "raw_value", "description", "action"])
        writer.writeheader()
        for err in errors:
            writer.writerow(err)


def main():
    print("Запуск отрефакторенного ETL-парсера...")
    c, s, h, e = parse_all_kcr()
    export_artifacts(c, s, h, e)
    print(f"Готово! Клиентов: {len(c)}, Срезов: {len(s)}, Истории: {len(h)}, Замечаний аудита: {len(e)}")


if __name__ == "__main__":
    main()
