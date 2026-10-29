"""
parser_core.py — ETL-модуль разбора Excel и нормализации в единый JSON-контракт.
Контур КЦР. Задача 1.
"""

from __future__ import annotations
import os
import sys
import json
import csv
from dataclasses import dataclass, asdict
from datetime import datetime, date
from typing import Any, Dict, List, Optional, Tuple, Set
import openpyxl


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
    time_zone_offset: Optional[str]
    admin_name: Optional[str]
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
    rejection_reason: Optional[str] = None

@dataclass
class HistoryRecord:
    source_client_id: str
    event_type: str
    occurred_at: Optional[str]
    text: Optional[str]
    raw_status: Optional[str]
    next_contact_date: Optional[str]
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
# Вспомогательные функции валидации и преобразования
# =====================================================================

def parse_iso_date(val: Any) -> Optional[str]:
    if val is None:
        return None
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")
    val_str = str(val).strip()
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d.%m.%y"):
        try:
            return datetime.strptime(val_str, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return None

def parse_strict_int(val: Any) -> Tuple[Optional[int], Optional[str]]:
    """Возвращает (число, ошибка). Не округляет дробные и не заменяет текст на null."""
    if val is None or str(val).strip() == "":
        return None, None
    if isinstance(val, bool):
        return None, "BOOLEAN_IN_NUMERIC_FIELD"
    if isinstance(val, int):
        return val, None
    if isinstance(val, float):
        if val.is_integer():
            return int(val), None
        return None, f"FRACTIONAL_NUMBER:{val}"
    val_str = str(val).strip().replace(" ", "").replace(",", ".")
    try:
        f_val = float(val_str)
        if f_val.is_integer():
            return int(f_val), None
        return None, f"FRACTIONAL_NUMBER:{val}"
    except ValueError:
        return None, f"INVALID_NUMERIC_STRING:{val}"

def detect_header_row(ws: openpyxl.worksheet.worksheet.Worksheet, max_scan: int = 5) -> int:
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
# Парсеры отдельных файлов
# =====================================================================

class BaseRegistryParser:
    SLICE_SPECS = [
        # весь 2025
        (15, 16, 17, "2025-01-01", "2025-12-31", "2026-09-29"),
        (15, 18, 19, "2025-01-01", "2025-12-31", "2026-07-21"),
        (15, 20, 21, "2025-01-01", "2025-12-31", "2026-07-07"),
        (15, 22, 23, "2025-01-01", "2025-12-31", "2026-05-21"),
        # 2-4 кв 2025
        (24, 25, 26, "2025-04-01", "2025-12-31", "2026-07-21"),
        (24, 27, 28, "2025-04-01", "2025-12-31", "2026-07-07"),
        (24, 29, 30, "2025-04-01", "2025-12-31", "2026-05-21"),
        # 1 кв 2026
        (31, 32, 33, "2026-01-01", "2026-03-31", "2026-09-29"),
        (31, 34, 35, "2026-01-01", "2026-03-31", "2026-07-21"),
        (31, 36, 37, "2026-01-01", "2026-03-31", "2026-07-07"),
        (31, 38, 39, "2026-01-01", "2026-03-31", "2026-05-21"),
        # 2 кв 2025 - 1 кв 2026
        (40, 41, 42, "2025-04-01", "2026-03-31", "2026-07-21"),
        (40, 43, 44, "2025-04-01", "2026-03-31", "2026-07-07"),
        (40, 45, 46, "2025-04-01", "2026-03-31", "2026-05-21"),
        # 3-4 кв 2025
        (47, 48, 49, "2025-07-01", "2025-12-31", "2026-09-29"),
        (47, 50, 51, "2025-07-01", "2025-12-31", "2026-07-21"),
        # 2 кв 2026
        (52, 53, 54, "2026-04-01", "2026-06-30", "2026-09-29"),
        (52, 55, 56, "2026-04-01", "2026-06-30", "2026-07-21"),
        # 3 кв 2025 - 2 кв 2026
        (57, 58, 59, "2025-07-01", "2026-06-30", "2026-09-29"),
        (57, 60, 61, "2025-07-01", "2026-06-30", "2026-07-21"),
        # весь 2025 - 2 кв 2026
        (62, 63, 64, "2025-01-01", "2026-06-30", "2026-09-29"),
        (62, 65, 66, "2025-01-01", "2026-06-30", "2026-07-21"),
    ]

    @classmethod
    def parse(cls, filepath: str) -> Tuple[Dict[str, ClientRecord], List[SliceRecord], List[QualityIssue]]:
        filename = os.path.basename(filepath)
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb["Все КЦР"]
        h_row = detect_header_row(ws)

        val_header = str(ws.cell(h_row, 15).value or "")
        if "Всего документов" not in val_header:
            raise ValueError(f"Структура {filename} нарушена: ожидался заголовок 'Всего документов' в столбце 15")

        clients: Dict[str, ClientRecord] = {}
        slices: List[SliceRecord] = []
        issues: List[QualityIssue] = []

        for r in range(h_row + 1, ws.max_row + 1):
            raw_id = ws.cell(r, 1).value
            if raw_id is None or str(raw_id).strip() == "":
                if any(ws.cell(r, c).value is not None for c in range(1, ws.max_column + 1)):
                    issues.append(QualityIssue(
                        code="MISSING_CLIENT_ID", file=filename, sheet="Все КЦР", row=r, col="A",
                        raw_value=None, description="Строка с данными без идентификатора клиента", action="skipped"
                    ))
                continue

            cid = str(raw_id).strip()
            if cid in clients:
                issues.append(QualityIssue(
                    code="DUPLICATE_CLIENT_ID", file=filename, sheet="Все КЦР", row=r, col="A",
                    raw_value=cid, description=f"Дубликат идентификатора клиента: {cid}", action="requires_review"
                ))
            else:
                clients[cid] = ClientRecord(
                    source_client_id=cid,
                    is_active=str(ws.cell(r, 10).value).strip() if ws.cell(r, 10).value else None,
                    is_blocked=str(ws.cell(r, 11).value).strip() if ws.cell(r, 11).value else None,
                    time_zone_offset=None,
                    admin_name=None,
                    provenance=Provenance(source_file=filename, source_sheet="Все КЦР", source_row=r)
                )

            for c_tot, c_uns, c_pct, p_start, p_end, as_of in cls.SLICE_SPECS:
                raw_tot = ws.cell(r, c_tot).value
                raw_uns = ws.cell(r, c_uns).value
                raw_pct = ws.cell(r, c_pct).value

                if raw_tot is None and raw_uns is None:
                    continue

                tot, err_tot = parse_strict_int(raw_tot)
                uns, err_uns = parse_strict_int(raw_uns)
                pct = float(raw_pct) if isinstance(raw_pct, (int, float)) else None

                col_name = f"{openpyxl.utils.get_column_letter(c_tot)},{openpyxl.utils.get_column_letter(c_uns)}"
                prov = Provenance(source_file=filename, source_sheet="Все КЦР", source_row=r, source_column=col_name)

                rejection = None
                if err_tot:
                    issues.append(QualityIssue(code="INVALID_NUMERIC_TOTAL", file=filename, sheet="Все КЦР", row=r, col=openpyxl.utils.get_column_letter(c_tot), raw_value=str(raw_tot), description=f"Некорректное значение total: {err_tot}", action="requires_review"))
                    rejection = err_tot
                if err_uns:
                    issues.append(QualityIssue(code="INVALID_NUMERIC_UNSENT", file=filename, sheet="Все КЦР", row=r, col=openpyxl.utils.get_column_letter(c_uns), raw_value=str(raw_uns), description=f"Некорректное значение unsent: {err_uns}", action="requires_review"))
                    rejection = err_uns

                slices.append(SliceRecord(
                    source_client_id=cid, period_start=p_start, period_end=p_end, as_of_date=as_of,
                    total=tot, unsent=uns, raw_percentage=pct, provenance=prov, rejection_reason=rejection
                ))

        return clients, slices, issues


class PeriodicSliceParser:
    SHEET_PERIODS = {
        "весь 2025 год": ("2025-01-01", "2025-12-31"),
        "1 кв 2026": ("2026-01-01", "2026-03-31"),
        "2 кв 2026": ("2026-04-01", "2026-06-30"),
        "3 кв 2026": ("2026-07-01", "2026-09-30")
    }

    @classmethod
    def parse(cls, filepath: str, known_clients: Set[str]) -> Tuple[List[SliceRecord], List[QualityIssue]]:
        filename = os.path.basename(filepath)
        wb = openpyxl.load_workbook(filepath, data_only=True)
        slices: List[SliceRecord] = []
        issues: List[QualityIssue] = []

        for sname, (p_start, p_end) in cls.SHEET_PERIODS.items():
            if sname not in wb.sheetnames:
                continue
            ws = wb[sname]
            h_row = detect_header_row(ws)
            seen_in_sheet: Set[str] = set()

            for r in range(h_row + 1, ws.max_row + 1):
                raw_id = ws.cell(r, 1).value
                if raw_id is None or str(raw_id).strip() == "":
                    continue

                cid = str(raw_id).strip()
                rejection = None

                if cid in seen_in_sheet:
                    issues.append(QualityIssue(
                        code="DUPLICATE_CLIENT_ID", file=filename, sheet=sname, row=r, col="A",
                        raw_value=cid, description=f"Дублирующийся идентификатор на листе: {cid}", action="requires_review"
                    ))
                    rejection = "DUPLICATE_CLIENT_ID"
                else:
                    seen_in_sheet.add(cid)

                if cid not in known_clients:
                    issues.append(QualityIssue(
                        code="UNKNOWN_CLIENT_ID", file=filename, sheet=sname, row=r, col="A",
                        raw_value=cid, description=f"Клиент {cid} отсутствует в основной базе", action="blocked"
                    ))
                    rejection = "UNKNOWN_CLIENT_ID"

                raw_tot = ws.cell(r, 4).value
                raw_uns = ws.cell(r, 5).value

                tot, err_tot = parse_strict_int(raw_tot)
                uns, err_uns = parse_strict_int(raw_uns)

                if err_tot:
                    issues.append(QualityIssue(code="INVALID_NUMERIC_TOTAL", file=filename, sheet=sname, row=r, col="D", raw_value=str(raw_tot), description=err_tot, action="requires_review"))
                    rejection = rejection or err_tot
                if err_uns:
                    issues.append(QualityIssue(code="INVALID_NUMERIC_UNSENT", file=filename, sheet=sname, row=r, col="E", raw_value=str(raw_uns), description=err_uns, action="requires_review"))
                    rejection = rejection or err_uns

                slices.append(SliceRecord(
                    source_client_id=cid, period_start=p_start, period_end=p_end, as_of_date="2026-10-20",
                    total=tot, unsent=uns, raw_percentage=None,
                    provenance=Provenance(source_file=filename, source_sheet=sname, source_row=r, source_column="D,E"),
                    rejection_reason=rejection
                ))

        return slices, issues


class HistoricalFilesParser:
    @classmethod
    def parse_debts_2025(cls, filepath: str) -> Tuple[List[SliceRecord], List[HistoryRecord], List[QualityIssue]]:
        filename = os.path.basename(filepath)
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb["1-3 кв 25"]
        h_row = detect_header_row(ws)

        slices: List[SliceRecord] = []
        history: List[HistoryRecord] = []
        issues: List[QualityIssue] = []

        issues.append(QualityIssue(
            code="HEADER_CONTRADICTION", file=filename, sheet="1-3 кв 25", row=h_row, col="W vs X",
            raw_value="W='долг на 16.01.25', X='% Долга на 16.01.26'",
            description="Противоречие дат в заголовках: столбец W озаглавлен 16.01.25, а связанный столбец X — 16.01.26",
            action="requires_review"
        ))

        hist_slice_specs = [
            (19, 35, None, "2025-10-20"),
            (19, 33, 34, "2025-10-31"),
            (19, 31, 32, "2025-11-05"),
            (19, 27, 28, "2025-12-05"),
            (19, 23, 24, "2026-01-16"),
            (19, 20, 21, "2026-03-27"),
        ]

        for r in range(h_row + 1, ws.max_row + 1):
            raw_id = ws.cell(r, 1).value
            if raw_id is None or str(raw_id).strip() == "":
                if any(ws.cell(r, c).value is not None for c in range(1, ws.max_column + 1)):
                    issues.append(QualityIssue(
                        code="MISSING_CLIENT_ID", file=filename, sheet="1-3 кв 25", row=r, col="A",
                        raw_value=None, description="Строка с данными без идентификатора клиента", action="skipped"
                    ))
                continue

            cid = str(raw_id).strip()

            for c_tot, c_uns, c_pct, as_of in hist_slice_specs:
                raw_tot = ws.cell(r, c_tot).value
                raw_uns = ws.cell(r, c_uns).value
                raw_pct = ws.cell(r, c_pct).value if c_pct else None

                if raw_tot is None and raw_uns is None:
                    continue

                tot, err_tot = parse_strict_int(raw_tot)
                uns, err_uns = parse_strict_int(raw_uns)
                pct = float(raw_pct) if isinstance(raw_pct, (int, float)) else None

                col_name = f"{openpyxl.utils.get_column_letter(c_tot)},{openpyxl.utils.get_column_letter(c_uns)}"
                prov = Provenance(source_file=filename, source_sheet="1-3 кв 25", source_row=r, source_column=col_name)

                rejection = err_tot or err_uns
                slices.append(SliceRecord(
                    source_client_id=cid, period_start="2025-01-01", period_end="2025-12-31", as_of_date=as_of,
                    total=tot, unsent=uns, raw_percentage=pct, provenance=prov, rejection_reason=rejection
                ))

            comm1 = ws.cell(r, 46).value
            stat1 = ws.cell(r, 47).value
            recal1 = parse_iso_date(ws.cell(r, 48).value)
            dl1 = ws.cell(r, 49).value
            client_date1 = ws.cell(r, 50).value
            deadline_combined = str(dl1).strip() if dl1 else (str(client_date1).strip() if client_date1 else None)

            if comm1 or stat1 or recal1 or deadline_combined:
                history.append(HistoryRecord(
                    source_client_id=cid, event_type="first_contact", occurred_at=None,
                    text=str(comm1).strip() if comm1 else None,
                    raw_status=str(stat1).strip() if stat1 else None,
                    next_contact_date=recal1, raw_deadline=deadline_combined,
                    provenance=Provenance(source_file=filename, source_sheet="1-3 кв 25", source_row=r, source_column="AT-AX")
                ))

            comm2 = ws.cell(r, 53).value
            stat2 = ws.cell(r, 54).value
            recal2 = parse_iso_date(ws.cell(r, 55).value)
            dl_by2 = ws.cell(r, 56).value
            dl_from2 = ws.cell(r, 57).value
            dl2_combined = f"{dl_by2}; {dl_from2}".strip("; ") if (dl_by2 or dl_from2) else None

            if comm2 or stat2 or recal2 or dl2_combined:
                history.append(HistoryRecord(
                    source_client_id=cid, event_type="second_contact", occurred_at=None,
                    text=str(comm2).strip() if comm2 else None,
                    raw_status=str(stat2).strip() if stat2 else None,
                    next_contact_date=recal2, raw_deadline=dl2_combined,
                    provenance=Provenance(source_file=filename, source_sheet="1-3 кв 25", source_row=r, source_column="BA-BE")
                ))

        return slices, history, issues

    @classmethod
    def parse_q1_2026(cls, filepath: str) -> Tuple[List[SliceRecord], List[HistoryRecord], List[QualityIssue]]:
        filename = os.path.basename(filepath)
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb["Все КЦР"]
        h_row = detect_header_row(ws)

        slices: List[SliceRecord] = []
        history: List[HistoryRecord] = []
        issues: List[QualityIssue] = []

        q1_slice_specs = [
            (22, 23, 24, "2025-01-01", "2025-12-31"),
            (25, 26, 27, "2025-04-01", "2025-12-31"),
            (28, 29, 30, "2026-01-01", "2026-03-31"),
            (31, 32, 33, "2025-04-01", "2026-03-31"),
        ]

        for r in range(h_row + 1, ws.max_row + 1):
            raw_id = ws.cell(r, 1).value
            if raw_id is None or str(raw_id).strip() == "":
                continue

            cid = str(raw_id).strip()

            for c_tot, c_uns, c_pct, p_start, p_end in q1_slice_specs:
                raw_tot = ws.cell(r, c_tot).value
                raw_uns = ws.cell(r, c_uns).value
                raw_pct = ws.cell(r, c_pct).value

                if raw_tot is None and raw_uns is None:
                    continue

                tot, err_tot = parse_strict_int(raw_tot)
                uns, err_uns = parse_strict_int(raw_uns)
                pct = float(raw_pct) if isinstance(raw_pct, (int, float)) else None

                col_name = f"{openpyxl.utils.get_column_letter(c_tot)},{openpyxl.utils.get_column_letter(c_uns)}"
                prov = Provenance(source_file=filename, source_sheet="Все КЦР", source_row=r, source_column=col_name)

                rejection = err_tot or err_uns
                slices.append(SliceRecord(
                    source_client_id=cid, period_start=p_start, period_end=p_end, as_of_date="2026-05-21",
                    total=tot, unsent=uns, raw_percentage=pct, provenance=prov, rejection_reason=rejection
                ))

            comm = ws.cell(r, 35).value
            stat = ws.cell(r, 36).value
            recal = parse_iso_date(ws.cell(r, 37).value)
            dl_client = ws.cell(r, 38).value
            dl_date = ws.cell(r, 39).value
            dl_combined = str(dl_client).strip() if dl_client else (str(dl_date).strip() if dl_date else None)

            if comm or stat or recal or dl_combined:
                history.append(HistoryRecord(
                    source_client_id=cid, event_type="operator_contact", occurred_at=None,
                    text=str(comm).strip() if comm else None,
                    raw_status=str(stat).strip() if stat else None,
                    next_contact_date=recal, raw_deadline=dl_combined,
                    provenance=Provenance(source_file=filename, source_sheet="Все КЦР", source_row=r, source_column="AI-AM")
                ))

        return slices, history, issues


# =====================================================================
# Централизованный аудит качества
# =====================================================================

class QualityAuditor:
    @staticmethod
    def audit_and_filter_slices(slices: List[SliceRecord]) -> Tuple[List[SliceRecord], List[SliceRecord], List[QualityIssue]]:
        valid: List[SliceRecord] = []
        rejected: List[SliceRecord] = []
        issues: List[QualityIssue] = []

        seen_keys: Dict[Tuple[str, str, str, str], SliceRecord] = {}

        for s in slices:
            p = s.provenance
            key = (s.source_client_id, s.period_start, s.period_end, s.as_of_date)

            if key in seen_keys:
                issues.append(QualityIssue(
                    code="DUPLICATE_SLICE", file=p.source_file, sheet=p.source_sheet, row=p.source_row,
                    col=p.source_column or "A", raw_value=str(key),
                    description=f"Дублирующийся срез данных для ключа {key}", action="requires_review"
                ))
                s.rejection_reason = "DUPLICATE_SLICE"
                rejected.append(s)
                continue
            seen_keys[key] = s

            if s.total is None and s.unsent is not None:
                issues.append(QualityIssue(
                    code="UNSENT_WITHOUT_TOTAL", file=p.source_file, sheet=p.source_sheet, row=p.source_row,
                    col=p.source_column or "total", raw_value=f"total=null, unsent={s.unsent}",
                    description="Указан остаток непереданных документов при отсутствующем общем количестве",
                    action="requires_review"
                ))
                s.rejection_reason = "UNSENT_WITHOUT_TOTAL"
                rejected.append(s)
                continue

            if s.total is not None and s.unsent is not None and s.unsent > s.total:
                issues.append(QualityIssue(
                    code="UNSENT_EXCEEDS_TOTAL", file=p.source_file, sheet=p.source_sheet, row=p.source_row,
                    col=p.source_column or "unsent", raw_value=f"unsent={s.unsent}, total={s.total}",
                    description="Количество непереданных документов превышает общее число документов",
                    action="requires_review"
                ))
                s.rejection_reason = "UNSENT_EXCEEDS_TOTAL"
                rejected.append(s)
                continue

            if s.raw_percentage is not None and s.total and s.unsent is not None and s.total > 0:
                calc_pct = s.unsent / s.total
                if abs(s.raw_percentage - calc_pct) > 0.02:
                    issues.append(QualityIssue(
                        code="PERCENTAGE_MISMATCH", file=p.source_file, sheet=p.source_sheet, row=p.source_row,
                        col=p.source_column or "percentage", raw_value=f"excel={s.raw_percentage:.4f}, calc={calc_pct:.4f}",
                        description="Исходный процент расходится с вычисленным unsent/total более чем на 0.02",
                        action="requires_review"
                    ))

            if s.rejection_reason:
                rejected.append(s)
            else:
                valid.append(s)

        return valid, rejected, issues


# =====================================================================
# Оркестрация и экспорт
# =====================================================================

def parse_all_kcr(base_dir: str = ".") -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], List[Dict[str, Any]], List[Dict[str, Any]], List[Dict[str, Any]]]:
    required_files = [
        "КЦР_последняя выгрузка_Дубна.xlsx",
        "Срез 20.10.26_Дубна.xlsx",
        "База должники III кв 2025 Дубна.xlsx",
        "База I кв 2026 Дубна.xlsx"
    ]
    missing = [f for f in required_files if not os.path.exists(os.path.join(base_dir, f))]
    if missing:
        raise FileNotFoundError(f"Отсутствуют обязательные входные файлы в '{base_dir}': {missing}")

    f1 = os.path.join(base_dir, "КЦР_последняя выгрузка_Дубна.xlsx")
    f2 = os.path.join(base_dir, "Срез 20.10.26_Дубна.xlsx")
    f3 = os.path.join(base_dir, "База должники III кв 2025 Дубна.xlsx")
    f4 = os.path.join(base_dir, "База I кв 2026 Дубна.xlsx")

    clients_dict, s1, issues1 = BaseRegistryParser.parse(f1)
    known_cids = set(clients_dict.keys())

    s2, issues2 = PeriodicSliceParser.parse(f2, known_cids)
    s3, h3, issues3 = HistoricalFilesParser.parse_debts_2025(f3)
    s4, h4, issues4 = HistoricalFilesParser.parse_q1_2026(f4)

    all_raw_slices = s1 + s2 + s3 + s4
    valid_slices, rejected_slices, audit_issues = QualityAuditor.audit_and_filter_slices(all_raw_slices)

    all_history = h3 + h4
    all_issues = issues1 + issues2 + issues3 + issues4 + audit_issues

    return (
        [asdict(c) for c in clients_dict.values()],
        [asdict(s) for s in valid_slices],
        [asdict(s) for s in rejected_slices],
        [asdict(h) for h in all_history],
        [asdict(i) for i in all_issues]
    )

def export_artifacts(clients, valid_slices, rejected_slices, history, issues, output_dir: str = "."):
    os.makedirs(output_dir, exist_ok=True)
    staging_dir = os.path.join(output_dir, "_staging")
    os.makedirs(staging_dir, exist_ok=True)

    with open(os.path.join(staging_dir, "normalized_clients.json"), "w", encoding="utf-8") as f:
        json.dump(clients, f, ensure_ascii=False, indent=2)

    with open(os.path.join(staging_dir, "normalized_slices.jsonl"), "w", encoding="utf-8") as f:
        for s in valid_slices:
            f.write(json.dumps(s, ensure_ascii=False) + "\n")

    with open(os.path.join(staging_dir, "rejected_slices.jsonl"), "w", encoding="utf-8") as f:
        for s in rejected_slices:
            f.write(json.dumps(s, ensure_ascii=False) + "\n")

    with open(os.path.join(staging_dir, "normalized_history.jsonl"), "w", encoding="utf-8") as f:
        for h in history:
            f.write(json.dumps(h, ensure_ascii=False) + "\n")

    with open(os.path.join(staging_dir, "quality_report.json"), "w", encoding="utf-8") as f:
        json.dump(issues, f, ensure_ascii=False, indent=2)

    with open(os.path.join(staging_dir, "quality_report.csv"), "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["code", "file", "sheet", "row", "col", "raw_value", "description", "action"])
        writer.writeheader()
        for i in issues:
            writer.writerow(i)

    for fname in os.listdir(staging_dir):
        src = os.path.join(staging_dir, fname)
        dst = os.path.join(output_dir, fname)
        if os.path.exists(dst):
            os.remove(dst)
        os.rename(src, dst)
    os.rmdir(staging_dir)

def main():
    base_dir = "." if os.path.exists("КЦР_последняя выгрузка_Дубна.xlsx") else ".."
    output_dir = os.path.dirname(os.path.abspath(__file__))

    print(f"Запуск ETL-модуля (входы: {base_dir}, выходы: {output_dir})...")
    c, vs, rs, h, issues = parse_all_kcr(base_dir=base_dir)
    export_artifacts(c, vs, rs, h, issues, output_dir=output_dir)

    print("\nСводка обработки:")
    print(f"  Клиентов в базе: {len(c)}")
    print(f"  Валидных срезов (допущены): {len(vs)}")
    print(f"  Отклонённых срезов (rejected): {len(rs)}")
    print(f"  Исторических контактов: {len(h)}")
    print(f"  Замечаний аудита качества: {len(issues)}")

if __name__ == "__main__":
    main()
