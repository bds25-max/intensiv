"""
parser_core.py — ETL-модуль разбора Excel и нормализации в единый JSON-контракт.
Контур КЦР. Задача 1.
Без маскирования и искусственного обезличивания: все тексты, комментарии и атрибуты сохраняются дословно.
"""

from __future__ import annotations
import os
import sys
import json
import csv
import re
import math
import shutil
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
    raw_next_contact: Optional[str] = None
    raw_send_deadline: Optional[str] = None
    raw_client_date: Optional[str] = None


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

def clean_optional_str(val: Any) -> Optional[str]:
    """Преобразует непустое значение в строку без изменений содержимого."""
    if val is None:
        return None
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    s = str(val).strip()
    return s if s != "" else None


def parse_iso_date(val: Any) -> Optional[str]:
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, date):
        return val.strftime("%Y-%m-%d")
    val_str = str(val).strip()
    if not val_str:
        return None
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d.%m.%y"):
        try:
            dt = datetime.strptime(val_str, fmt)
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            pass
    return None


def is_valid_iso_date(val: Any) -> bool:
    if not isinstance(val, str) or len(val) != 10:
        return False
    try:
        datetime.strptime(val, "%Y-%m-%d")
        return True
    except ValueError:
        return False


def parse_strict_int(val: Any) -> Tuple[Optional[int], Optional[str]]:
    """Возвращает (число, ошибка). Отклоняет NaN/Inf, дробные и нечисловые строки."""
    if val is None or str(val).strip() == "":
        return None, None
    if isinstance(val, bool):
        return None, f"BOOLEAN_IN_NUMERIC_FIELD:{val}"
    if isinstance(val, int):
        return val, None
    if isinstance(val, float):
        if not math.isfinite(val):
            return None, f"NON_FINITE_NUMBER:{val}"
        if val.is_integer():
            return int(val), None
        return None, f"FRACTIONAL_NUMBER:{val}"
    val_str = str(val).strip().replace(" ", "").replace(",", ".")
    try:
        f_val = float(val_str)
        if not math.isfinite(f_val):
            return None, f"NON_FINITE_NUMBER:{val}"
        if f_val.is_integer():
            return int(f_val), None
        return None, f"FRACTIONAL_NUMBER:{val}"
    except ValueError:
        return None, f"INVALID_NUMERIC_STRING:{val}"


def parse_strict_float(val: Any) -> Tuple[Optional[float], Optional[str]]:
    """Строгий разбор вещественного числа (процента). Отклоняет NaN, Infinity, #N/A и текст."""
    if val is None or str(val).strip() == "":
        return None, None
    if isinstance(val, bool):
        return None, f"BOOLEAN_IN_FLOAT_FIELD:{val}"
    if isinstance(val, (int, float)):
        f_val = float(val)
        if not math.isfinite(f_val):
            return None, f"NON_FINITE_FLOAT:{val}"
        return f_val, None
    val_str = str(val).strip().replace(" ", "").replace(",", ".")
    if val_str.endswith("%"):
        try:
            f_val = float(val_str[:-1]) / 100.0
            if not math.isfinite(f_val):
                return None, f"NON_FINITE_FLOAT:{val}"
            return f_val, None
        except ValueError:
            return None, f"INVALID_FLOAT_STRING:{val}"
    try:
        f_val = float(val_str)
        if not math.isfinite(f_val):
            return None, f"NON_FINITE_FLOAT:{val}"
        return f_val, None
    except ValueError:
        return None, f"INVALID_FLOAT_STRING:{val}"


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


def normalize_header_for_match(text: str) -> str:
    """Нормализует пробелы и дефисы для точной сверки периода и даты в заголовке."""
    s = (text or "").strip().lower()
    s = re.sub(r"\s*-\s*", "-", s)
    s = re.sub(r"\s+", " ", s)
    return s


def verify_header_tokens(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    h_row: int,
    expected_tokens: Dict[int, List[str]],
    filename: str,
    sheetname: str
) -> None:
    """Проверяет, что в заголовках указанных колонок присутствуют все ожидаемые смысловые фрагменты (включая период и дату)."""
    for col_idx, tokens in expected_tokens.items():
        raw_h = normalize_header_for_match(str(ws.cell(h_row, col_idx).value or ""))
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        for tok in tokens:
            norm_tok = normalize_header_for_match(tok)
            if norm_tok not in raw_h:
                raise ValueError(
                    f"Структура {filename} (лист '{sheetname}') нарушена в колонке {col_letter} (столбец {col_idx}): "
                    f"ожидался фрагмент '{tok}', фактически '{ws.cell(h_row, col_idx).value}'"
                )


def extract_dates_from_header(header_text: str) -> List[str]:
    """Извлекает даты вида ДД.ММ.ГГ или ДД.ММ.ГГГГ из заголовка и возвращает в ISO."""
    matches = re.findall(r"\b(\d{1,2}\.\d{2}\.\d{2,4})\b", header_text or "")
    res = []
    for m in matches:
        iso = parse_iso_date(m)
        if iso:
            res.append(iso)
    return res


def parse_slice_cell_triplet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    r: int,
    c_tot: int,
    c_uns: int,
    c_pct: Optional[int],
    cid: str,
    p_start: str,
    p_end: str,
    as_of: str,
    filename: str,
    sheetname: str,
    issues: List[QualityIssue],
    base_rejection: Optional[str] = None
) -> Optional[SliceRecord]:
    """Единая точка извлечения и первичной числовой проверки тройки ячеек (total, unsent, raw_percentage)."""
    raw_tot = ws.cell(r, c_tot).value
    raw_uns = ws.cell(r, c_uns).value
    raw_pct = ws.cell(r, c_pct).value if c_pct else None

    if raw_tot is None and raw_uns is None:
        return None

    tot, err_tot = parse_strict_int(raw_tot)
    uns, err_uns = parse_strict_int(raw_uns)
    pct, err_pct = parse_strict_float(raw_pct) if c_pct else (None, None)

    col_name = f"{openpyxl.utils.get_column_letter(c_tot)},{openpyxl.utils.get_column_letter(c_uns)}"
    prov = Provenance(source_file=filename, source_sheet=sheetname, source_row=r, source_column=col_name)

    rejection = base_rejection
    if err_tot:
        issues.append(QualityIssue(
            code="INVALID_NUMERIC_TOTAL", file=filename, sheet=sheetname, row=r,
            col=openpyxl.utils.get_column_letter(c_tot), raw_value=str(raw_tot),
            description=f"Некорректное значение total: {err_tot}", action="requires_review"
        ))
        rejection = rejection or err_tot
    if err_uns:
        issues.append(QualityIssue(
            code="INVALID_NUMERIC_UNSENT", file=filename, sheet=sheetname, row=r,
            col=openpyxl.utils.get_column_letter(c_uns), raw_value=str(raw_uns),
            description=f"Некорректное значение unsent: {err_uns}", action="requires_review"
        ))
        rejection = rejection or err_uns
    if err_pct:
        issues.append(QualityIssue(
            code="INVALID_NUMERIC_PERCENTAGE", file=filename, sheet=sheetname, row=r,
            col=openpyxl.utils.get_column_letter(c_pct), raw_value=str(raw_pct),
            description=f"Некорректное значение процента: {err_pct}", action="requires_review"
        ))
        rejection = rejection or err_pct

    return SliceRecord(
        source_client_id=cid,
        period_start=p_start,
        period_end=p_end,
        as_of_date=as_of,
        total=tot,
        unsent=uns,
        raw_percentage=pct,
        provenance=prov,
        rejection_reason=rejection
    )


# =====================================================================
# Парсеры отдельных файлов
# =====================================================================

class BaseRegistryParser:
    # (col_total, col_unsent, col_pct, period_start, period_end, as_of_date, period_token, date_token)
    SLICE_SPECS = [
        # весь 2025
        (15, 16, 17, "2025-01-01", "2025-12-31", "2026-09-29", "весь 2025", "29.09.26"),
        (15, 18, 19, "2025-01-01", "2025-12-31", "2026-07-21", "весь 2025", "21.07.26"),
        (15, 20, 21, "2025-01-01", "2025-12-31", "2026-07-07", "весь 2025", "07.07.26"),
        (15, 22, 23, "2025-01-01", "2025-12-31", "2026-05-21", "весь 2025", "21.05.26"),
        # 2-4 кв 2025
        (24, 25, 26, "2025-04-01", "2025-12-31", "2026-07-21", "2-4", "21.07.26"),
        (24, 27, 28, "2025-04-01", "2025-12-31", "2026-07-07", "2-4", "07.07.26"),
        (24, 29, 30, "2025-04-01", "2025-12-31", "2026-05-21", "2-4", "21.05.26"),
        # 1 кв 2026
        (31, 32, 33, "2026-01-01", "2026-03-31", "2026-09-29", "1 кв 2026", "29.09.26"),
        (31, 34, 35, "2026-01-01", "2026-03-31", "2026-07-21", "1 кв 2026", "21.07.26"),
        (31, 36, 37, "2026-01-01", "2026-03-31", "2026-07-07", "1 кв 2026", "07.07.26"),
        (31, 38, 39, "2026-01-01", "2026-03-31", "2026-05-21", "1 кв 2026", "21.05.26"),
        # 2 кв 2025 - 1 кв 2026
        (40, 41, 42, "2025-04-01", "2026-03-31", "2026-07-21", "2 кв 2025-1 кв 2026", "21.07.26"),
        (40, 43, 44, "2025-04-01", "2026-03-31", "2026-07-07", "2 кв 2025-1 кв 2026", "07.07.26"),
        (40, 45, 46, "2025-04-01", "2026-03-31", "2026-05-21", "2 кв 2025-1 кв 2026", "21.05.26"),
        # 3-4 кв 2025
        (47, 48, 49, "2025-07-01", "2025-12-31", "2026-09-29", "3-4 кв 2025", "29.09.26"),
        (47, 50, 51, "2025-07-01", "2025-12-31", "2026-07-21", "3-4 кв 2025", "21.07.26"),
        # 2 кв 2026
        (52, 53, 54, "2026-04-01", "2026-06-30", "2026-09-29", "2 кв 2026", "29.09.26"),
        (52, 55, 56, "2026-04-01", "2026-06-30", "2026-07-21", "2 кв 2026", "21.07.26"),
        # 3 кв 2025 - 2 кв 2026
        (57, 58, 59, "2025-07-01", "2026-06-30", "2026-09-29", "3 кв 2025-2 кв 2026", "29.09.26"),
        (57, 60, 61, "2025-07-01", "2026-06-30", "2026-07-21", "3 кв 2025-2 кв 2026", "21.07.26"),
        # весь 2025 - 2 кв 2026
        (62, 63, 64, "2025-01-01", "2026-06-30", "2026-09-29", "весь 2025-2 кв 2026", "29.09.26"),
        (62, 65, 66, "2025-01-01", "2026-06-30", "2026-07-21", "весь 2025-2 кв 2026", "21.07.26"),
    ]

    @classmethod
    def parse(cls, filepath: str) -> Tuple[Dict[str, ClientRecord], List[SliceRecord], List[QualityIssue]]:
        filename = os.path.basename(filepath)
        wb = openpyxl.load_workbook(filepath, data_only=True)
        try:
            if "Все КЦР" not in wb.sheetnames:
                raise ValueError(f"В файле {filename} отсутствует обязательный лист 'Все КЦР'")
            ws = wb["Все КЦР"]
            h_row = detect_header_row(ws)

            expected_headers: Dict[int, List[str]] = {
                10: ["действует"],
                11: ["блокировка"],
            }
            for c_tot, c_uns, c_pct, _, _, _, p_tok, d_tok in cls.SLICE_SPECS:
                expected_headers[c_tot] = ["всего", p_tok]
                expected_headers[c_uns] = ["долг", p_tok, d_tok]
                expected_headers[c_pct] = ["%", p_tok, d_tok]

            verify_header_tokens(ws, h_row, expected_headers, filename, "Все КЦР")

            clients: Dict[str, ClientRecord] = {}
            slices: List[SliceRecord] = []
            issues: List[QualityIssue] = []

            for r in range(h_row + 1, ws.max_row + 1):
                raw_id = ws.cell(r, 1).value
                if raw_id is None or str(raw_id).strip() == "":
                    if any(ws.cell(r, c).value is not None for c in range(1, ws.max_column + 1)):
                        issues.append(QualityIssue(
                            code="MISSING_CLIENT_ID", file=filename, sheet="Все КЦР", row=r, col="A",
                            raw_value=None if raw_id is None else repr(raw_id),
                            description="Строка с данными без идентификатора клиента", action="skipped"
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
                        is_active=clean_optional_str(ws.cell(r, 10).value),
                        is_blocked=clean_optional_str(ws.cell(r, 11).value),
                        time_zone_offset=None,
                        admin_name=clean_optional_str(ws.cell(r, 12).value) or clean_optional_str(ws.cell(r, 13).value),
                        provenance=Provenance(source_file=filename, source_sheet="Все КЦР", source_row=r)
                    )

                for c_tot, c_uns, c_pct, p_start, p_end, as_of, _, _ in cls.SLICE_SPECS:
                    rec = parse_slice_cell_triplet(
                        ws, r, c_tot, c_uns, c_pct, cid, p_start, p_end, as_of, filename, "Все КЦР", issues
                    )
                    if rec is not None:
                        slices.append(rec)

            return clients, slices, issues
        finally:
            wb.close()


class PeriodicSliceParser:
    SHEET_PERIODS = {
        "весь 2025 год": ("2025-01-01", "2025-12-31"),
        "1 кв 2026": ("2026-01-01", "2026-03-31"),
        "2 кв 2026": ("2026-04-01", "2026-06-30"),
        "3 кв 2026": ("2026-07-01", "2026-09-30")
    }

    @classmethod
    def parse(cls, filepath: str, known_clients: Set[str], require_all_sheets: bool = False) -> Tuple[List[SliceRecord], List[QualityIssue]]:
        filename = os.path.basename(filepath)
        wb = openpyxl.load_workbook(filepath, data_only=True)
        try:
            slices: List[SliceRecord] = []
            issues: List[QualityIssue] = []

            if require_all_sheets:
                missing_sheets = [s for s in cls.SHEET_PERIODS if s not in wb.sheetnames]
                if missing_sheets:
                    raise ValueError(f"В файле {filename} отсутствуют обязательные листы: {missing_sheets}")

            for sname, (p_start, p_end) in cls.SHEET_PERIODS.items():
                if sname not in wb.sheetnames:
                    issues.append(QualityIssue(
                        code="MISSING_SHEET", file=filename, sheet=sname, row=0, col="A",
                        raw_value=sname, description=f"Отсутствует ожидаемый лист '{sname}'", action="requires_review"
                    ))
                    continue

                ws = wb[sname]
                h_row = detect_header_row(ws)
                col1_h = str(ws.cell(h_row, 1).value or "").lower()
                id_token = ["id"] if "id" in col1_h else ["идентификатор"]
                verify_header_tokens(ws, h_row, {
                    1: id_token,
                    4: ["всего"],
                    5: ["не отправлено"]
                }, filename, sname)

                for r in range(h_row + 1, ws.max_row + 1):
                    raw_id = ws.cell(r, 1).value
                    if raw_id is None or str(raw_id).strip() == "":
                        if any(ws.cell(r, c).value is not None for c in range(1, ws.max_column + 1)):
                            issues.append(QualityIssue(
                                code="MISSING_CLIENT_ID", file=filename, sheet=sname, row=r, col="A",
                                raw_value=None if raw_id is None else repr(raw_id),
                                description="Строка с данными без идентификатора клиента", action="skipped"
                            ))
                        continue

                    cid = str(raw_id).strip()
                    base_rejection = None

                    if cid not in known_clients:
                        issues.append(QualityIssue(
                            code="UNKNOWN_CLIENT_ID", file=filename, sheet=sname, row=r, col="A",
                            raw_value=cid, description=f"Клиент {cid} отсутствует в основной базе", action="blocked"
                        ))
                        base_rejection = "UNKNOWN_CLIENT_ID"

                    rec = parse_slice_cell_triplet(
                        ws, r, 4, 5, None, cid, p_start, p_end, "2026-10-20", filename, sname, issues, base_rejection=base_rejection
                    )
                    if rec is not None:
                        slices.append(rec)

            return slices, issues
        finally:
            wb.close()


class HistoricalFilesParser:
    # Столбцы исторических статусов K–Q (11–17) и AO–AP (41–42)
    STATUS_COLUMNS_SPEC = [
        (11, "26.05.26", "2026-05-26"),
        (12, "27.03.26", "2026-03-27"),
        (13, "13.03.26", "2026-03-13"),
        (14, "02.03.26", "2026-03-02"),
        (15, "16.01.26", "2026-01-16"),
        (16, "29.12.25", "2025-12-29"),
        (17, "5.12.25", "2025-12-05"),
        (41, "20.10.25", "2025-10-20"),
        (42, "05.11.25", "2025-11-05"),
    ]

    @classmethod
    def parse_debts_2025(
        cls,
        filepath: str,
        known_clients: Optional[Set[str]] = None,
        clients_dict: Optional[Dict[str, ClientRecord]] = None
    ) -> Tuple[List[SliceRecord], List[HistoryRecord], List[QualityIssue]]:
        filename = os.path.basename(filepath)
        wb = openpyxl.load_workbook(filepath, data_only=True)
        try:
            if "1-3 кв 25" not in wb.sheetnames:
                raise ValueError(f"В файле {filename} отсутствует обязательный лист '1-3 кв 25'")
            ws = wb["1-3 кв 25"]
            h_row = detect_header_row(ws)

            expected_headers: Dict[int, List[str]] = {
                1: ["идентификатор"],
                10: ["претензии"],
                19: ["документов всего", "2025"],
                20: ["долг", "27.03.26"],
                27: ["долг", "05.12.25"],
                31: ["долг", "05.11.25"],
                33: ["долг", "31.10.25"],
                35: ["долг", "20.10.25"],
                46: ["комментарии оовс"],
                47: ["статус звонка"],
                53: ["повторному звонку"],
                54: ["статус звонка в январе"],
            }
            for col_idx, date_tok, _ in cls.STATUS_COLUMNS_SPEC:
                expected_headers[col_idx] = ["статус", date_tok]

            verify_header_tokens(ws, h_row, expected_headers, filename, "1-3 кв 25")

            slices: List[SliceRecord] = []
            history: List[HistoryRecord] = []
            issues: List[QualityIssue] = []

            # Динамическая проверка заголовков W (23) и X (24) на противоречие дат
            w_header = str(ws.cell(h_row, 23).value or "").strip()
            x_header = str(ws.cell(h_row, 24).value or "").strip()
            w_dates = extract_dates_from_header(w_header)
            x_dates = extract_dates_from_header(x_header)

            w_x_ambiguous = False
            if not w_dates or not x_dates or w_dates[0] != x_dates[0]:
                w_x_ambiguous = True
                w_d_str = w_dates[0] if w_dates else "unknown"
                x_d_str = x_dates[0] if x_dates else "unknown"
                issues.append(QualityIssue(
                    code="HEADER_CONTRADICTION", file=filename, sheet="1-3 кв 25", row=h_row, col="W vs X",
                    raw_value=f"W={w_header!r}, X={x_header!r}",
                    description=f"Противоречие дат в заголовках столбцов W ({w_d_str}) и X ({x_d_str})",
                    action="requires_review"
                ))
                w_x_as_of = x_dates[0] if x_dates else (w_dates[0] if w_dates else "2026-01-16")
            else:
                w_x_as_of = w_dates[0]

            hist_slice_specs = [
                (19, 35, None, "2025-10-20", False),
                (19, 33, 34, "2025-10-31", False),
                (19, 31, 32, "2025-11-05", False),
                (19, 27, 28, "2025-12-05", False),
                (19, 23, 24, w_x_as_of, w_x_ambiguous),
                (19, 20, 21, "2026-03-27", False),
            ]

            for r in range(h_row + 1, ws.max_row + 1):
                raw_id = ws.cell(r, 1).value
                if raw_id is None or str(raw_id).strip() == "":
                    if any(ws.cell(r, c).value is not None for c in range(1, ws.max_column + 1)):
                        issues.append(QualityIssue(
                            code="MISSING_CLIENT_ID", file=filename, sheet="1-3 кв 25", row=r, col="A",
                            raw_value=None if raw_id is None else repr(raw_id),
                            description="Строка с данными без идентификатора клиента", action="skipped"
                        ))
                    continue

                cid = str(raw_id).strip()
                is_unknown = known_clients is not None and cid not in known_clients
                if is_unknown:
                    issues.append(QualityIssue(
                        code="UNKNOWN_CLIENT_ID", file=filename, sheet="1-3 кв 25", row=r, col="A",
                        raw_value=cid, description=f"Клиент {cid} отсутствует в основной базе", action="blocked"
                    ))
                elif clients_dict is not None and cid in clients_dict:
                    # Обогащаем карточку клиента доступными полями (Время от Екб, ФИО админа) без какого-либо обезличивания
                    tz_val = clean_optional_str(ws.cell(r, 5).value)
                    adm_val = clean_optional_str(ws.cell(r, 6).value)
                    if tz_val and not clients_dict[cid].time_zone_offset:
                        clients_dict[cid].time_zone_offset = tz_val
                    if adm_val and not clients_dict[cid].admin_name:
                        clients_dict[cid].admin_name = adm_val

                # 1. Числовые срезы
                for c_tot, c_uns, c_pct, as_of, is_ambig in hist_slice_specs:
                    base_rej = "UNKNOWN_CLIENT_ID" if is_unknown else (
                        "AMBIGUOUS_AS_OF_DATE:2025-01-16_vs_2026-01-16" if is_ambig else None
                    )
                    rec = parse_slice_cell_triplet(
                        ws, r, c_tot, c_uns, c_pct, cid, "2025-01-01", "2025-12-31", as_of,
                        filename, "1-3 кв 25", issues, base_rejection=base_rej
                    )
                    if rec is not None:
                        slices.append(rec)

                if is_unknown:
                    continue

                # 2. Исторические статусы K–Q (11–17) и AO–AP (41–42) — сохраняются дословно без маскирования
                for col_idx, _, iso_dt in cls.STATUS_COLUMNS_SPEC:
                    stat_val = clean_optional_str(ws.cell(r, col_idx).value)
                    if stat_val:
                        col_letter = openpyxl.utils.get_column_letter(col_idx)
                        history.append(HistoryRecord(
                            source_client_id=cid,
                            event_type="historical_status",
                            occurred_at=iso_dt,
                            text=None,
                            raw_status=stat_val,
                            next_contact_date=None,
                            raw_deadline=None,
                            provenance=Provenance(source_file=filename, source_sheet="1-3 кв 25", source_row=r, source_column=col_letter)
                        ))

                # 3. Комментарий для претензии (столбец J / 10) — дословно
                claim_comm = clean_optional_str(ws.cell(r, 10).value)
                if claim_comm:
                    history.append(HistoryRecord(
                        source_client_id=cid, event_type="claim_comment", occurred_at=None,
                        text=claim_comm, raw_status=None, next_contact_date=None, raw_deadline=None,
                        provenance=Provenance(source_file=filename, source_sheet="1-3 кв 25", source_row=r, source_column="J")
                    ))

                # 4. Первое касание (AT-AX: 46-50) — все телефоны, email и ФИО сохраняются в первозданном виде
                comm1 = clean_optional_str(ws.cell(r, 46).value)
                stat1 = clean_optional_str(ws.cell(r, 47).value)
                raw_recal1 = clean_optional_str(ws.cell(r, 48).value)
                recal1 = parse_iso_date(ws.cell(r, 48).value)
                if raw_recal1 and not recal1:
                    issues.append(QualityIssue(
                        code="NON_DATE_NEXT_CONTACT", file=filename, sheet="1-3 кв 25", row=r, col="AV",
                        raw_value=raw_recal1, description="Некалендарное значение в поле 'Дата перезвона'", action="requires_review"
                    ))
                dl_send1 = clean_optional_str(ws.cell(r, 49).value)
                dl_client1 = clean_optional_str(ws.cell(r, 50).value)
                dl1_parts = [p for p in (dl_send1, dl_client1) if p is not None]
                dl1_combined = "; ".join(dl1_parts) if dl1_parts else None

                if any([comm1, stat1, raw_recal1, dl_send1, dl_client1]):
                    history.append(HistoryRecord(
                        source_client_id=cid, event_type="first_contact", occurred_at=None,
                        text=comm1, raw_status=stat1, next_contact_date=recal1, raw_deadline=dl1_combined,
                        provenance=Provenance(source_file=filename, source_sheet="1-3 кв 25", source_row=r, source_column="AT-AX"),
                        raw_next_contact=raw_recal1, raw_send_deadline=dl_send1, raw_client_date=dl_client1
                    ))

                # 5. Повторный контакт в январе (BA-BE: 53-57) — дословно
                comm2 = clean_optional_str(ws.cell(r, 53).value)
                stat2 = clean_optional_str(ws.cell(r, 54).value)
                raw_recal2 = clean_optional_str(ws.cell(r, 55).value)
                recal2 = parse_iso_date(ws.cell(r, 55).value)
                if raw_recal2 and not recal2:
                    issues.append(QualityIssue(
                        code="NON_DATE_NEXT_CONTACT", file=filename, sheet="1-3 кв 25", row=r, col="BC",
                        raw_value=raw_recal2, description="Некалендарное значение в поле 'Дата перезвона'", action="requires_review"
                    ))
                dl_send2 = clean_optional_str(ws.cell(r, 56).value)
                dl_client2 = clean_optional_str(ws.cell(r, 57).value)
                dl2_parts = [p for p in (dl_send2, dl_client2) if p is not None]
                dl2_combined = "; ".join(dl2_parts) if dl2_parts else None

                if any([comm2, stat2, raw_recal2, dl_send2, dl_client2]):
                    history.append(HistoryRecord(
                        source_client_id=cid, event_type="second_contact", occurred_at=None,
                        text=comm2, raw_status=stat2, next_contact_date=recal2, raw_deadline=dl2_combined,
                        provenance=Provenance(source_file=filename, source_sheet="1-3 кв 25", source_row=r, source_column="BA-BE"),
                        raw_next_contact=raw_recal2, raw_send_deadline=dl_send2, raw_client_date=dl_client2
                    ))

                # 6. Инциденты, инфо по документам и запрос эксперта (BF-BH: 58-60) — дословно
                doc_info = clean_optional_str(ws.cell(r, 58).value)
                incident_no = clean_optional_str(ws.cell(r, 59).value)
                expert_req = clean_optional_str(ws.cell(r, 60).value)
                if any([doc_info, incident_no, expert_req]):
                    meta_parts = []
                    if doc_info:
                        meta_parts.append(f"Инфо от клиента: {doc_info}")
                    if incident_no:
                        meta_parts.append(f"Инцидент: {incident_no}")
                    if expert_req:
                        meta_parts.append(f"Запрос эксперта: {expert_req}")
                    history.append(HistoryRecord(
                        source_client_id=cid, event_type="incident_or_expert_request", occurred_at=None,
                        text="; ".join(meta_parts), raw_status=expert_req, next_contact_date=None, raw_deadline=doc_info,
                        provenance=Provenance(source_file=filename, source_sheet="1-3 кв 25", source_row=r, source_column="BF-BH")
                    ))

            return slices, history, issues
        finally:
            wb.close()

    @classmethod
    def parse_q1_2026(
        cls,
        filepath: str,
        known_clients: Optional[Set[str]] = None,
        clients_dict: Optional[Dict[str, ClientRecord]] = None
    ) -> Tuple[List[SliceRecord], List[HistoryRecord], List[QualityIssue]]:
        filename = os.path.basename(filepath)
        wb = openpyxl.load_workbook(filepath, data_only=True)
        try:
            if "Все КЦР" not in wb.sheetnames:
                raise ValueError(f"В файле {filename} отсутствует обязательный лист 'Все КЦР'")
            ws = wb["Все КЦР"]
            h_row = detect_header_row(ws)

            verify_header_tokens(ws, h_row, {
                1: ["идентификатор"],
                22: ["всего", "весь 2025"],
                23: ["долг", "весь 2025", "21.05.26"],
                24: ["%", "весь 2025", "21.05.26"],
                25: ["всего", "2-4"],
                26: ["долг", "2-4", "21.05.26"],
                27: ["%", "2-4", "21.05.26"],
                28: ["всего", "1 кв 2026"],
                29: ["долг", "1 кв 2026", "21.05.26"],
                30: ["%", "1 кв 2026", "21.05.26"],
                31: ["всего", "2 кв 2025-1 кв 2026"],
                32: ["долг", "2 кв 2025-1 кв 2026", "21.05.26"],
                33: ["%", "2 кв 2025-1 кв 2026", "21.05.26"],
                35: ["комментарии оовс"],
                36: ["статус звонка"],
            }, filename, "Все КЦР")

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
                    if any(ws.cell(r, c).value is not None for c in range(1, ws.max_column + 1)):
                        issues.append(QualityIssue(
                            code="MISSING_CLIENT_ID", file=filename, sheet="Все КЦР", row=r, col="A",
                            raw_value=None if raw_id is None else repr(raw_id),
                            description="Строка с данными без идентификатора клиента", action="skipped"
                        ))
                    continue

                cid = str(raw_id).strip()
                is_unknown = known_clients is not None and cid not in known_clients
                if is_unknown:
                    issues.append(QualityIssue(
                        code="UNKNOWN_CLIENT_ID", file=filename, sheet="Все КЦР", row=r, col="A",
                        raw_value=cid, description=f"Клиент {cid} отсутствует в основной базе", action="blocked"
                    ))
                elif clients_dict is not None and cid in clients_dict:
                    # Обогащаем карточку клиента часовым поясом (столбец J / 10: 'Время от Екб') и ФИО админа (K / 11)
                    tz_val = clean_optional_str(ws.cell(r, 10).value)
                    adm_val = clean_optional_str(ws.cell(r, 11).value)
                    if tz_val and not clients_dict[cid].time_zone_offset:
                        clients_dict[cid].time_zone_offset = tz_val
                    if adm_val and not clients_dict[cid].admin_name:
                        clients_dict[cid].admin_name = adm_val

                for c_tot, c_uns, c_pct, p_start, p_end in q1_slice_specs:
                    base_rej = "UNKNOWN_CLIENT_ID" if is_unknown else None
                    rec = parse_slice_cell_triplet(
                        ws, r, c_tot, c_uns, c_pct, cid, p_start, p_end, "2026-05-21",
                        filename, "Все КЦР", issues, base_rejection=base_rej
                    )
                    if rec is not None:
                        slices.append(rec)

                if is_unknown:
                    continue

                comm = clean_optional_str(ws.cell(r, 35).value)
                stat = clean_optional_str(ws.cell(r, 36).value)
                raw_recal = clean_optional_str(ws.cell(r, 37).value)
                recal = parse_iso_date(ws.cell(r, 37).value)
                if raw_recal and not recal:
                    issues.append(QualityIssue(
                        code="NON_DATE_NEXT_CONTACT", file=filename, sheet="Все КЦР", row=r, col="AK",
                        raw_value=raw_recal, description="Некалендарное значение в поле 'Дата перезвона'", action="requires_review"
                    ))
                dl_send = clean_optional_str(ws.cell(r, 38).value)
                dl_client = clean_optional_str(ws.cell(r, 39).value)
                dl_parts = [p for p in (dl_send, dl_client) if p is not None]
                dl_combined = "; ".join(dl_parts) if dl_parts else None

                if any([comm, stat, raw_recal, dl_send, dl_client]):
                    history.append(HistoryRecord(
                        source_client_id=cid, event_type="operator_contact", occurred_at=None,
                        text=comm, raw_status=stat, next_contact_date=recal, raw_deadline=dl_combined,
                        provenance=Provenance(source_file=filename, source_sheet="Все КЦР", source_row=r, source_column="AI-AM"),
                        raw_next_contact=raw_recal, raw_send_deadline=dl_send, raw_client_date=dl_client
                    ))

            return slices, history, issues
        finally:
            wb.close()


# =====================================================================
# Централизованный аудит качества
# =====================================================================

class QualityAuditor:
    @staticmethod
    def audit_and_filter_slices(slices: List[SliceRecord]) -> Tuple[List[SliceRecord], List[SliceRecord], List[QualityIssue]]:
        valid_by_key: Dict[Tuple[str, str, str, str], SliceRecord] = {}
        conflicted_keys: Set[Tuple[str, str, str, str]] = set()
        rejected: List[SliceRecord] = []
        issues: List[QualityIssue] = []

        for s in slices:
            p = s.provenance

            # 1. Сначала проверяем индивидуальную пригодность записи (до регистрации ключа!)
            if not (is_valid_iso_date(s.period_start) and is_valid_iso_date(s.period_end) and is_valid_iso_date(s.as_of_date)):
                issues.append(QualityIssue(
                    code="INVALID_DATE_VALUE", file=p.source_file, sheet=p.source_sheet, row=p.source_row,
                    col=p.source_column or "date",
                    raw_value=f"start={s.period_start}, end={s.period_end}, as_of={s.as_of_date}",
                    description="Некорректная или пустая календарная дата периода или среза",
                    action="requires_review"
                ))
                s.rejection_reason = s.rejection_reason or "INVALID_DATE_VALUE"
                rejected.append(s)
                continue

            if s.period_start > s.period_end:
                issues.append(QualityIssue(
                    code="INVALID_PERIOD_BOUNDS", file=p.source_file, sheet=p.source_sheet, row=p.source_row,
                    col=p.source_column or "period",
                    raw_value=f"{s.period_start} > {s.period_end}",
                    description="Начало отчётного периода позже его окончания",
                    action="requires_review"
                ))
                s.rejection_reason = s.rejection_reason or "INVALID_PERIOD_BOUNDS"
                rejected.append(s)
                continue

            if s.raw_percentage is not None and not math.isfinite(s.raw_percentage):
                issues.append(QualityIssue(
                    code="INVALID_NUMERIC_PERCENTAGE", file=p.source_file, sheet=p.source_sheet, row=p.source_row,
                    col=p.source_column or "percentage", raw_value=str(s.raw_percentage),
                    description="Неконечное значение процента (NaN или Infinity)", action="requires_review"
                ))
                s.rejection_reason = s.rejection_reason or f"NON_FINITE_FLOAT:{s.raw_percentage}"

            if s.total is not None and s.total < 0:
                issues.append(QualityIssue(
                    code="NEGATIVE_TOTAL", file=p.source_file, sheet=p.source_sheet, row=p.source_row,
                    col=p.source_column or "total", raw_value=str(s.total),
                    description="Отрицательное общее количество документов", action="requires_review"
                ))
                s.rejection_reason = s.rejection_reason or "NEGATIVE_TOTAL"

            if s.unsent is not None and s.unsent < 0:
                issues.append(QualityIssue(
                    code="NEGATIVE_UNSENT", file=p.source_file, sheet=p.source_sheet, row=p.source_row,
                    col=p.source_column or "unsent", raw_value=str(s.unsent),
                    description="Отрицательное количество непереданных документов", action="requires_review"
                ))
                s.rejection_reason = s.rejection_reason or "NEGATIVE_UNSENT"

            if s.total is None and s.unsent is not None and not s.rejection_reason:
                issues.append(QualityIssue(
                    code="UNSENT_WITHOUT_TOTAL", file=p.source_file, sheet=p.source_sheet, row=p.source_row,
                    col=p.source_column or "total", raw_value=f"total=null, unsent={s.unsent}",
                    description="Указан остаток непереданных документов при отсутствующем общем количестве",
                    action="requires_review"
                ))
                s.rejection_reason = "UNSENT_WITHOUT_TOTAL"

            if s.total is not None and s.unsent is not None and s.unsent > s.total and not s.rejection_reason:
                issues.append(QualityIssue(
                    code="UNSENT_EXCEEDS_TOTAL", file=p.source_file, sheet=p.source_sheet, row=p.source_row,
                    col=p.source_column or "unsent", raw_value=f"unsent={s.unsent}, total={s.total}",
                    description="Количество непереданных документов превышает общее число документов",
                    action="requires_review"
                ))
                s.rejection_reason = "UNSENT_EXCEEDS_TOTAL"

            if s.raw_percentage is not None and math.isfinite(s.raw_percentage) and s.total and s.unsent is not None and s.total > 0:
                calc_pct = s.unsent / s.total
                if abs(s.raw_percentage - calc_pct) > 0.02:
                    issues.append(QualityIssue(
                        code="PERCENTAGE_MISMATCH", file=p.source_file, sheet=p.source_sheet, row=p.source_row,
                        col=p.source_column or "percentage", raw_value=f"excel={s.raw_percentage:.4f}, calc={calc_pct:.4f}",
                        description="Исходный процент расходится с вычисленным unsent/total более чем на 0.02",
                        action="requires_review"
                    ))

            # Если запись уже отклонена по любой причине — она НЕ занимает ключ
            if s.rejection_reason:
                rejected.append(s)
                continue

            # 2. Обработка совпадений и конфликтов среди пригодных записей (включая повторы внутри одного листа!)
            key = (s.source_client_id, s.period_start, s.period_end, s.as_of_date)
            if key in conflicted_keys:
                issues.append(QualityIssue(
                    code="CONFLICTING_DUPLICATE_SLICE", file=p.source_file, sheet=p.source_sheet, row=p.source_row,
                    col=p.source_column or "A", raw_value=str(key),
                    description=f"Повторная запись для конфликтующего ключа {key}", action="requires_review"
                ))
                s.rejection_reason = "CONFLICTING_DUPLICATE_SLICE"
                rejected.append(s)
                continue

            if key in valid_by_key:
                prev = valid_by_key[key]
                if (prev.total, prev.unsent) != (s.total, s.unsent):
                    # Конфликт значений между двумя валидными записями: отклоняем обе
                    valid_by_key.pop(key)
                    conflicted_keys.add(key)
                    prev.rejection_reason = "CONFLICTING_DUPLICATE_SLICE"
                    s.rejection_reason = "CONFLICTING_DUPLICATE_SLICE"
                    rejected.append(prev)
                    rejected.append(s)
                    issues.append(QualityIssue(
                        code="CONFLICTING_DUPLICATE_SLICE", file=p.source_file, sheet=p.source_sheet, row=p.source_row,
                        col=p.source_column or "A",
                        raw_value=f"prev=({prev.total},{prev.unsent}) vs cur=({s.total},{s.unsent})",
                        description=f"Конфликт значений дублирующегося среза для ключа {key}", action="requires_review"
                    ))
                else:
                    issues.append(QualityIssue(
                        code="DUPLICATE_SLICE", file=p.source_file, sheet=p.source_sheet, row=p.source_row,
                        col=p.source_column or "A", raw_value=str(key),
                        description=f"Дублирующийся срез данных с идентичными значениями для ключа {key}",
                        action="requires_review"
                    ))
                    s.rejection_reason = "DUPLICATE_SLICE"
                    rejected.append(s)
                continue

            valid_by_key[key] = s

        return list(valid_by_key.values()), rejected, issues


# =====================================================================
# Оркестрация и транзакционный экспорт
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

    s2, issues2 = PeriodicSliceParser.parse(f2, known_cids, require_all_sheets=True)
    s3, h3, issues3 = HistoricalFilesParser.parse_debts_2025(f3, known_clients=known_cids, clients_dict=clients_dict)
    s4, h4, issues4 = HistoricalFilesParser.parse_q1_2026(f4, known_clients=known_cids, clients_dict=clients_dict)

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


def export_artifacts(clients, valid_slices, rejected_slices, history, issues, output_dir: str = ".") -> None:
    os.makedirs(output_dir, exist_ok=True)
    staging_dir = os.path.join(output_dir, "_staging")
    backup_dir = os.path.join(output_dir, "_backup")

    target_files = [
        "normalized_clients.json",
        "normalized_slices.jsonl",
        "rejected_slices.jsonl",
        "normalized_history.jsonl",
        "quality_report.json",
        "quality_report.csv"
    ]

    # Если с предыдущего прерванного запуска осталась непустая папка _backup,
    # НЕ удаляем её до восстановления недостающих целевых файлов!
    if os.path.exists(backup_dir):
        for fname in target_files:
            bck_file = os.path.join(backup_dir, fname)
            dst_file = os.path.join(output_dir, fname)
            if os.path.exists(bck_file) and not os.path.exists(dst_file):
                os.replace(bck_file, dst_file)

    if os.path.exists(staging_dir):
        shutil.rmtree(staging_dir)
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

    os.makedirs(backup_dir, exist_ok=True)
    moved_to_backup: List[str] = []
    published: List[str] = []
    commit_succeeded = False
    rollback_succeeded = False

    try:
        for fname in target_files:
            dst = os.path.join(output_dir, fname)
            if os.path.exists(dst):
                os.replace(dst, os.path.join(backup_dir, fname))
                moved_to_backup.append(fname)

        for fname in target_files:
            src = os.path.join(staging_dir, fname)
            dst = os.path.join(output_dir, fname)
            os.replace(src, dst)
            published.append(fname)

        commit_succeeded = True
    except Exception:
        for fname in published:
            dst = os.path.join(output_dir, fname)
            if os.path.exists(dst):
                os.remove(dst)
        for fname in moved_to_backup:
            bck = os.path.join(backup_dir, fname)
            dst = os.path.join(output_dir, fname)
            if os.path.exists(bck):
                os.replace(bck, dst)
        rollback_succeeded = True
        raise
    finally:
        if os.path.exists(staging_dir):
            shutil.rmtree(staging_dir)
        if (commit_succeeded or rollback_succeeded) and os.path.exists(backup_dir):
            shutil.rmtree(backup_dir)


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
    print(f"  Исторических материалов: {len(h)}")
    print(f"  Замечаний аудита качества: {len(issues)}")


if __name__ == "__main__":
    main()
