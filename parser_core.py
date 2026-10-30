"""
parser_core.py — ETL-модуль разбора Excel и нормализации в единый JSON-контракт.
Контур КЦР. Задача 1 (Версия 8: полный прогон по четырём исходным Excel,
повторный срез без total опирается только на принятые наблюдения).
"""

from __future__ import annotations
import os
import sys
import json
import csv
import re
import math
import shutil
from decimal import Decimal, InvalidOperation
from dataclasses import dataclass, asdict, field
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
    kcr_code: Optional[str] = None
    organization_name: Optional[str] = None
    service_center: Optional[str] = None
    inn: Optional[str] = None
    kpp: Optional[str] = None
    creation_date: Optional[str] = None
    admin_phone: Optional[str] = None
    admin_email: Optional[str] = None
    manager_name: Optional[str] = None
    attribute_provenance: Dict[str, Provenance] = field(default_factory=dict)


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


# Полная программная карта всех 188 колонок 4 исходных файлов (7 листов)
COLUMN_MAPPING: Dict[str, List[Dict[str, Any]]] = {'КЦР_последняя выгрузка_Дубна.xlsx::Все КЦР': [{'col_letter': 'A',
                                                 'col_index': 1,
                                                 'header': None,
                                                 'target_field': 'ClientRecord.source_client_id',
                                                 'transform': 'str.strip(); пустые пропускаются, дубли фиксируются',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 5874},
                                                {'col_letter': 'B',
                                                 'col_index': 2,
                                                 'header': 'Код КЦР',
                                                 'target_field': 'ClientRecord.kcr_code',
                                                 'transform': 'clean_optional_str (в обезличенной выборке пусто)',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 0},
                                                {'col_letter': 'C',
                                                 'col_index': 3,
                                                 'header': 'Название',
                                                 'target_field': 'ClientRecord.organization_name',
                                                 'transform': 'clean_optional_str (в обезличенной выборке пусто)',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 0},
                                                {'col_letter': 'D',
                                                 'col_index': 4,
                                                 'header': 'СЦ',
                                                 'target_field': 'ClientRecord.service_center',
                                                 'transform': 'clean_optional_str (в обезличенной выборке пусто)',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 0},
                                                {'col_letter': 'E',
                                                 'col_index': 5,
                                                 'header': None,
                                                 'target_field': None,
                                                 'transform': 'Сохранение исходного значения и координат в '
                                                              'source_preserved_values.jsonl при наличии значения.',
                                                 'exclusion_reason': 'Нет непустых значений ниже заголовка в '
                                                                     'предоставленном файле.',
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 0},
                                                {'col_letter': 'F',
                                                 'col_index': 6,
                                                 'header': 'Статус на 21.07.26',
                                                 'target_field': 'HistoryRecord(historical_status)',
                                                 'transform': "occurred_at='2026-07-21', source_column='F' (7 записей)",
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 7},
                                                {'col_letter': 'G',
                                                 'col_index': 7,
                                                 'header': 'стаус на 21.05.26',
                                                 'target_field': 'HistoryRecord(historical_status)',
                                                 'transform': "occurred_at='2026-05-21', source_column='G' (2 748 "
                                                              'записей)',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 2748},
                                                {'col_letter': 'H',
                                                 'col_index': 8,
                                                 'header': 'ИНН',
                                                 'target_field': 'ClientRecord.inn',
                                                 'transform': 'clean_optional_str (в обезличенной выборке пусто)',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 0},
                                                {'col_letter': 'I',
                                                 'col_index': 9,
                                                 'header': 'КПП',
                                                 'target_field': 'ClientRecord.kpp',
                                                 'transform': 'clean_optional_str (в обезличенной выборке пусто)',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 0},
                                                {'col_letter': 'J',
                                                 'col_index': 10,
                                                 'header': 'Действует ли КЦР',
                                                 'target_field': 'ClientRecord.is_active',
                                                 'transform': 'clean_optional_str (5 873 значения)',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 5873},
                                                {'col_letter': 'K',
                                                 'col_index': 11,
                                                 'header': 'Блокировка в сервисе',
                                                 'target_field': 'ClientRecord.is_blocked',
                                                 'transform': 'clean_optional_str (5 873 значения)',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 5873},
                                                {'col_letter': 'L',
                                                 'col_index': 12,
                                                 'header': 'ФИО менеджера',
                                                 'target_field': 'ClientRecord.manager_name',
                                                 'transform': 'clean_optional_str; сохраняется отдельно от admin_name',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 0},
                                                {'col_letter': 'M',
                                                 'col_index': 13,
                                                 'header': 'ФИО менеджера из КП',
                                                 'target_field': 'ClientRecord.manager_name',
                                                 'transform': 'Резервное поле manager_name, если L пусто',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 0},
                                                {'col_letter': 'N',
                                                 'col_index': 14,
                                                 'header': 'Статус на 21.07.26',
                                                 'target_field': None,
                                                 'transform': 'Сохранение исходного значения и координат в '
                                                              'source_preserved_values.jsonl при наличии значения.',
                                                 'exclusion_reason': 'Нет непустых значений ниже заголовка в '
                                                                     'предоставленном файле.',
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 0},
                                                {'col_letter': 'O',
                                                 'col_index': 15,
                                                 'header': 'Всего документов весь 2025',
                                                 'target_field': 'SliceRecord.total',
                                                 'transform': 'parse_strict_int; период 2025-01-01..2025-12-31',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3419},
                                                {'col_letter': 'P',
                                                 'col_index': 16,
                                                 'header': 'Долг весь 2025 на 29.09.26',
                                                 'target_field': 'SliceRecord.unsent',
                                                 'transform': 'parse_strict_int; период 2025-01-01..2025-12-31, срез '
                                                              '2026-09-29',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3419},
                                                {'col_letter': 'Q',
                                                 'col_index': 17,
                                                 'header': '% Долга весь 2025 на 29.09.26',
                                                 'target_field': 'SliceRecord.raw_percentage',
                                                 'transform': 'parse_strict_float; сверка с unsent/total (±0.02), срез '
                                                              '2026-09-29',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3419},
                                                {'col_letter': 'R',
                                                 'col_index': 18,
                                                 'header': 'Долг весь 2025 на 21.07.26',
                                                 'target_field': 'SliceRecord.unsent',
                                                 'transform': 'parse_strict_int; период 2025-01-01..2025-12-31, срез '
                                                              '2026-07-21',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3419},
                                                {'col_letter': 'S',
                                                 'col_index': 19,
                                                 'header': '% Долга весь 2025 на 21.07.26',
                                                 'target_field': 'SliceRecord.raw_percentage',
                                                 'transform': 'parse_strict_float; сверка с unsent/total (±0.02), срез '
                                                              '2026-07-21',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3419},
                                                {'col_letter': 'T',
                                                 'col_index': 20,
                                                 'header': 'Долг весь 2025 на 07.07.26',
                                                 'target_field': 'SliceRecord.unsent',
                                                 'transform': 'parse_strict_int; период 2025-01-01..2025-12-31, срез '
                                                              '2026-07-07',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3419},
                                                {'col_letter': 'U',
                                                 'col_index': 21,
                                                 'header': '% Долга весь 2025 на 07.07.26',
                                                 'target_field': 'SliceRecord.raw_percentage',
                                                 'transform': 'parse_strict_float; сверка с unsent/total (±0.02), срез '
                                                              '2026-07-07',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3418},
                                                {'col_letter': 'V',
                                                 'col_index': 22,
                                                 'header': 'Долг весь 2025 на 21.05.26',
                                                 'target_field': 'SliceRecord.unsent',
                                                 'transform': 'parse_strict_int; период 2025-01-01..2025-12-31, срез '
                                                              '2026-05-21',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3418},
                                                {'col_letter': 'W',
                                                 'col_index': 23,
                                                 'header': '% Долга весь 2025 на 21.05.26',
                                                 'target_field': 'SliceRecord.raw_percentage',
                                                 'transform': 'parse_strict_float; сверка с unsent/total (±0.02), срез '
                                                              '2026-05-21',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3418},
                                                {'col_letter': 'X',
                                                 'col_index': 24,
                                                 'header': 'Всего документов 2-4 кв 2025',
                                                 'target_field': 'SliceRecord.total',
                                                 'transform': 'parse_strict_int; период 2025-04-01..2025-12-31',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3298},
                                                {'col_letter': 'Y',
                                                 'col_index': 25,
                                                 'header': 'Долг 2-4 кв 2025 на 21.07.26',
                                                 'target_field': 'SliceRecord.unsent',
                                                 'transform': 'parse_strict_int; период 2025-04-01..2025-12-31, срез '
                                                              '2026-07-21',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3298},
                                                {'col_letter': 'Z',
                                                 'col_index': 26,
                                                 'header': '% Долга 2-4кв 2025 на 21.07.26',
                                                 'target_field': 'SliceRecord.raw_percentage',
                                                 'transform': 'parse_strict_float; сверка с unsent/total (±0.02), срез '
                                                              '2026-07-21',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3298},
                                                {'col_letter': 'AA',
                                                 'col_index': 27,
                                                 'header': 'Долг 2-4 кв 2025 на 07.07.26',
                                                 'target_field': 'SliceRecord.unsent',
                                                 'transform': 'parse_strict_int; период 2025-04-01..2025-12-31, срез '
                                                              '2026-07-07',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3298},
                                                {'col_letter': 'AB',
                                                 'col_index': 28,
                                                 'header': '% Долга 2-4кв 2025 на 07.07.26',
                                                 'target_field': 'SliceRecord.raw_percentage',
                                                 'transform': 'parse_strict_float; сверка с unsent/total (±0.02), срез '
                                                              '2026-07-07',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3297},
                                                {'col_letter': 'AC',
                                                 'col_index': 29,
                                                 'header': 'Долг 2-4 кв 2025 на 21.05.26',
                                                 'target_field': 'SliceRecord.unsent',
                                                 'transform': 'parse_strict_int; период 2025-04-01..2025-12-31, срез '
                                                              '2026-05-21',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3297},
                                                {'col_letter': 'AD',
                                                 'col_index': 30,
                                                 'header': '% Долга 2-4кв 2025 на 21.05.26',
                                                 'target_field': 'SliceRecord.raw_percentage',
                                                 'transform': 'parse_strict_float; сверка с unsent/total (±0.02), срез '
                                                              '2026-05-21',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3297},
                                                {'col_letter': 'AE',
                                                 'col_index': 31,
                                                 'header': 'Всего документов 1 кв 2026',
                                                 'target_field': 'SliceRecord.total',
                                                 'transform': 'parse_strict_int; период 2026-01-01..2026-03-31',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3181},
                                                {'col_letter': 'AF',
                                                 'col_index': 32,
                                                 'header': 'Долг 1 кв 2026 на 29.09.26',
                                                 'target_field': 'SliceRecord.unsent',
                                                 'transform': 'parse_strict_int; период 2026-01-01..2026-03-31, срез '
                                                              '2026-09-29',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3181},
                                                {'col_letter': 'AG',
                                                 'col_index': 33,
                                                 'header': '% Долга 1 кв 2026 на 29.09.26',
                                                 'target_field': 'SliceRecord.raw_percentage',
                                                 'transform': 'parse_strict_float; сверка с unsent/total (±0.02), срез '
                                                              '2026-09-29',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3181},
                                                {'col_letter': 'AH',
                                                 'col_index': 34,
                                                 'header': 'Долг 1 кв 2026 на 21.07.26',
                                                 'target_field': 'SliceRecord.unsent',
                                                 'transform': 'parse_strict_int; период 2026-01-01..2026-03-31, срез '
                                                              '2026-07-21',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 2928},
                                                {'col_letter': 'AI',
                                                 'col_index': 35,
                                                 'header': '% Долга 1 кв 2026 на 21.07.26',
                                                 'target_field': 'SliceRecord.raw_percentage',
                                                 'transform': 'parse_strict_float; сверка с unsent/total (±0.02), срез '
                                                              '2026-07-21',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 55},
                                                {'col_letter': 'AJ',
                                                 'col_index': 36,
                                                 'header': 'Долг 1 кв 2026 на 07.07.26',
                                                 'target_field': 'SliceRecord.unsent',
                                                 'transform': 'parse_strict_int; период 2026-01-01..2026-03-31, срез '
                                                              '2026-07-07',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 2928},
                                                {'col_letter': 'AK',
                                                 'col_index': 37,
                                                 'header': '% Долга 1 кв 2026 на 07.07.26',
                                                 'target_field': 'SliceRecord.raw_percentage',
                                                 'transform': 'parse_strict_float; сверка с unsent/total (±0.02), срез '
                                                              '2026-07-07',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 2928},
                                                {'col_letter': 'AL',
                                                 'col_index': 38,
                                                 'header': 'Долг 1 кв 2026 на 21.05.26',
                                                 'target_field': 'SliceRecord.unsent',
                                                 'transform': 'parse_strict_int; период 2026-01-01..2026-03-31, срез '
                                                              '2026-05-21',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3180},
                                                {'col_letter': 'AM',
                                                 'col_index': 39,
                                                 'header': '% Долга 1 кв 2026 на 21.05.26',
                                                 'target_field': 'SliceRecord.raw_percentage',
                                                 'transform': 'parse_strict_float; сверка с unsent/total (±0.02), срез '
                                                              '2026-05-21',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3180},
                                                {'col_letter': 'AN',
                                                 'col_index': 40,
                                                 'header': 'Всего документов 2 кв 2025-1 кв 2026',
                                                 'target_field': 'SliceRecord.total',
                                                 'transform': 'parse_strict_int; период 2025-04-01..2026-03-31',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3657},
                                                {'col_letter': 'AO',
                                                 'col_index': 41,
                                                 'header': 'Долг 2 кв 2025-1 кв 2026 на 21.07.26',
                                                 'target_field': 'SliceRecord.unsent',
                                                 'transform': 'parse_strict_int; период 2025-04-01..2026-03-31, срез '
                                                              '2026-07-21',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3657},
                                                {'col_letter': 'AP',
                                                 'col_index': 42,
                                                 'header': '% Долга 2 кв 2025-1 кв 2026 на 21.07.26',
                                                 'target_field': 'SliceRecord.raw_percentage',
                                                 'transform': 'parse_strict_float; сверка с unsent/total (±0.02), срез '
                                                              '2026-07-21',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3657},
                                                {'col_letter': 'AQ',
                                                 'col_index': 43,
                                                 'header': 'Долг 2 кв 2025-1 кв 2026 на 07.07.26',
                                                 'target_field': 'SliceRecord.unsent',
                                                 'transform': 'parse_strict_int; период 2025-04-01..2026-03-31, срез '
                                                              '2026-07-07',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3657},
                                                {'col_letter': 'AR',
                                                 'col_index': 44,
                                                 'header': '% Долга 2 кв 2025-1 кв 2026 на 07.07.26',
                                                 'target_field': 'SliceRecord.raw_percentage',
                                                 'transform': 'parse_strict_float; сверка с unsent/total (±0.02), срез '
                                                              '2026-07-07',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3657},
                                                {'col_letter': 'AS',
                                                 'col_index': 45,
                                                 'header': 'Долг 2 кв 2025-1 кв 2026 на 21.05.26',
                                                 'target_field': 'SliceRecord.unsent',
                                                 'transform': 'parse_strict_int; период 2025-04-01..2026-03-31, срез '
                                                              '2026-05-21',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 5713},
                                                {'col_letter': 'AT',
                                                 'col_index': 46,
                                                 'header': '% Долга 2 кв 2025-1 кв 2026 на 21.05.26',
                                                 'target_field': 'SliceRecord.raw_percentage',
                                                 'transform': 'parse_strict_float; сверка с unsent/total (±0.02), срез '
                                                              '2026-05-21',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3656},
                                                {'col_letter': 'AU',
                                                 'col_index': 47,
                                                 'header': 'Всего документов 3-4 кв 2025',
                                                 'target_field': 'SliceRecord.total',
                                                 'transform': 'parse_strict_int; период 2025-07-01..2025-12-31',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3153},
                                                {'col_letter': 'AV',
                                                 'col_index': 48,
                                                 'header': 'Долг 3-4 кв 2025 на 29.09.26',
                                                 'target_field': 'SliceRecord.unsent',
                                                 'transform': 'parse_strict_int; период 2025-07-01..2025-12-31, срез '
                                                              '2026-09-29',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3298},
                                                {'col_letter': 'AW',
                                                 'col_index': 49,
                                                 'header': '% Долг 3-4 кв 2025 на 29.09.26',
                                                 'target_field': 'SliceRecord.raw_percentage',
                                                 'transform': 'parse_strict_float; сверка с unsent/total (±0.02), срез '
                                                              '2026-09-29',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3153},
                                                {'col_letter': 'AX',
                                                 'col_index': 50,
                                                 'header': 'Долг 3-4 кв 2025 на 21.07.26',
                                                 'target_field': 'SliceRecord.unsent',
                                                 'transform': 'parse_strict_int; период 2025-07-01..2025-12-31, срез '
                                                              '2026-07-21',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3153},
                                                {'col_letter': 'AY',
                                                 'col_index': 51,
                                                 'header': '% Долг 3-4 кв 2025 на 21.07.26',
                                                 'target_field': 'SliceRecord.raw_percentage',
                                                 'transform': 'parse_strict_float; сверка с unsent/total (±0.02), срез '
                                                              '2026-07-21',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3153},
                                                {'col_letter': 'AZ',
                                                 'col_index': 52,
                                                 'header': 'Всего документов 2 кв 2026',
                                                 'target_field': 'SliceRecord.total',
                                                 'transform': 'parse_strict_int; период 2026-04-01..2026-06-30',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3021},
                                                {'col_letter': 'BA',
                                                 'col_index': 53,
                                                 'header': 'Долг 2 кв 2026 на 29.09.26',
                                                 'target_field': 'SliceRecord.unsent',
                                                 'transform': 'parse_strict_int; период 2026-04-01..2026-06-30, срез '
                                                              '2026-09-29',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3021},
                                                {'col_letter': 'BB',
                                                 'col_index': 54,
                                                 'header': '% Долг 2 кв 2026 на 29.09.26',
                                                 'target_field': 'SliceRecord.raw_percentage',
                                                 'transform': 'parse_strict_float; сверка с unsent/total (±0.02), срез '
                                                              '2026-09-29',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3021},
                                                {'col_letter': 'BC',
                                                 'col_index': 55,
                                                 'header': 'Долг 2 кв 2026 на 21.07.26',
                                                 'target_field': 'SliceRecord.unsent',
                                                 'transform': 'parse_strict_int; период 2026-04-01..2026-06-30, срез '
                                                              '2026-07-21',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3021},
                                                {'col_letter': 'BD',
                                                 'col_index': 56,
                                                 'header': '% Долг 2 кв 2026 на 21.07.26',
                                                 'target_field': 'SliceRecord.raw_percentage',
                                                 'transform': 'parse_strict_float; сверка с unsent/total (±0.02), срез '
                                                              '2026-07-21',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 2591},
                                                {'col_letter': 'BE',
                                                 'col_index': 57,
                                                 'header': 'Всего документов 3 кв 2025 - 2 кв 2026',
                                                 'target_field': 'SliceRecord.total',
                                                 'transform': 'parse_strict_int; период 2025-07-01..2026-06-30',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3666},
                                                {'col_letter': 'BF',
                                                 'col_index': 58,
                                                 'header': 'Долг по документам 3 кв 2025-2 кв 2026 на 29.09.26',
                                                 'target_field': 'SliceRecord.unsent',
                                                 'transform': 'parse_strict_int; период 2025-07-01..2026-06-30, срез '
                                                              '2026-09-29',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3666},
                                                {'col_letter': 'BG',
                                                 'col_index': 59,
                                                 'header': '% Долг по документам 3 кв 2025-2 кв 2026 на 29.09.26',
                                                 'target_field': 'SliceRecord.raw_percentage',
                                                 'transform': 'parse_strict_float; сверка с unsent/total (±0.02), срез '
                                                              '2026-09-29',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3666},
                                                {'col_letter': 'BH',
                                                 'col_index': 60,
                                                 'header': 'Долг по документам 3 кв 2025-2 кв 2026 на 21.07.26',
                                                 'target_field': 'SliceRecord.unsent',
                                                 'transform': 'parse_strict_int; период 2025-07-01..2026-06-30, срез '
                                                              '2026-07-21',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3666},
                                                {'col_letter': 'BI',
                                                 'col_index': 61,
                                                 'header': '% Долг по документам 3 кв 2025-2 кв 2026 на 21.07.26',
                                                 'target_field': 'SliceRecord.raw_percentage',
                                                 'transform': 'parse_strict_float; сверка с unsent/total (±0.02), срез '
                                                              '2026-07-21',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3666},
                                                {'col_letter': 'BJ',
                                                 'col_index': 62,
                                                 'header': 'Всего документов весь 2025 - 2 кв 2026',
                                                 'target_field': 'SliceRecord.total',
                                                 'transform': 'parse_strict_int; период 2025-01-01..2026-06-30',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3846},
                                                {'col_letter': 'BK',
                                                 'col_index': 63,
                                                 'header': 'Долг по документам  весь 2025 - 2 кв 2026 на 29.09.26',
                                                 'target_field': 'SliceRecord.unsent',
                                                 'transform': 'parse_strict_int; период 2025-01-01..2026-06-30, срез '
                                                              '2026-09-29',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3846},
                                                {'col_letter': 'BL',
                                                 'col_index': 64,
                                                 'header': '% Долг по документам  весь 2025 - 2 кв 2026 на 29.09.26',
                                                 'target_field': 'SliceRecord.raw_percentage',
                                                 'transform': 'parse_strict_float; сверка с unsent/total (±0.02), срез '
                                                              '2026-09-29',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3846},
                                                {'col_letter': 'BM',
                                                 'col_index': 65,
                                                 'header': 'Долг по документам  весь 2025 - 2 кв 2026 на 21.07.26',
                                                 'target_field': 'SliceRecord.unsent',
                                                 'transform': 'parse_strict_int; период 2025-01-01..2026-06-30, срез '
                                                              '2026-07-21',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3846},
                                                {'col_letter': 'BN',
                                                 'col_index': 66,
                                                 'header': '% Долг по документам  весь 2025 - 2 кв 2026 на 21.07.26',
                                                 'target_field': 'SliceRecord.raw_percentage',
                                                 'transform': 'parse_strict_float; сверка с unsent/total (±0.02), срез '
                                                              '2026-07-21',
                                                 'exclusion_reason': None,
                                                 'header_row': 2,
                                                 'nonempty_cells_below_header': 3846}],
 'Срез 20.10.26_Дубна.xlsx::весь 2025 год': [{'col_letter': 'A',
                                              'col_index': 1,
                                              'header': 'Уникальны обезличинный идентификатор',
                                              'target_field': 'SliceRecord.source_client_id',
                                              'transform': 'Сопоставление с белым списком; неизвестные -> '
                                                           'UNKNOWN_CLIENT_ID (blocked)',
                                              'exclusion_reason': None,
                                              'header_row': 2,
                                              'nonempty_cells_below_header': 3419},
                                             {'col_letter': 'B',
                                              'col_index': 2,
                                              'header': 'Код КЦР',
                                              'target_field': 'ClientRecord.kcr_code',
                                              'transform': 'clean_optional_str (обогащение карточки клиента при '
                                                           'наличии)',
                                              'exclusion_reason': None,
                                              'header_row': 2,
                                              'nonempty_cells_below_header': 0},
                                             {'col_letter': 'C',
                                              'col_index': 3,
                                              'header': 'Наитменование организации',
                                              'target_field': 'ClientRecord.organization_name',
                                              'transform': 'clean_optional_str (обогащение карточки клиента при '
                                                           'наличии)',
                                              'exclusion_reason': None,
                                              'header_row': 2,
                                              'nonempty_cells_below_header': 0},
                                             {'col_letter': 'D',
                                              'col_index': 4,
                                              'header': 'документов всего',
                                              'target_field': 'SliceRecord.total',
                                              'transform': 'parse_strict_int; период 2025-01-01..2025-12-31, срез '
                                                           '2026-10-20',
                                              'exclusion_reason': None,
                                              'header_row': 2,
                                              'nonempty_cells_below_header': 3419},
                                             {'col_letter': 'E',
                                              'col_index': 5,
                                              'header': 'Всего не отправлено документов',
                                              'target_field': 'SliceRecord.unsent',
                                              'transform': 'parse_strict_int; период 2025-01-01..2025-12-31, срез '
                                                           '2026-10-20',
                                              'exclusion_reason': None,
                                              'header_row': 2,
                                              'nonempty_cells_below_header': 3419}],
 'Срез 20.10.26_Дубна.xlsx::1 кв 2026': [{'col_letter': 'A',
                                          'col_index': 1,
                                          'header': 'Уникальны обезличинный идентификатор',
                                          'target_field': 'SliceRecord.source_client_id',
                                          'transform': 'Сопоставление с белым списком; неизвестные -> '
                                                       'UNKNOWN_CLIENT_ID (blocked)',
                                          'exclusion_reason': None,
                                          'header_row': 2,
                                          'nonempty_cells_below_header': 2928},
                                         {'col_letter': 'B',
                                          'col_index': 2,
                                          'header': 'код СЦ',
                                          'target_field': 'ClientRecord.service_center',
                                          'transform': 'clean_optional_str (обогащение карточки клиента при наличии)',
                                          'exclusion_reason': None,
                                          'header_row': 2,
                                          'nonempty_cells_below_header': 0},
                                         {'col_letter': 'C',
                                          'col_index': 3,
                                          'header': 'название СЦ',
                                          'target_field': 'ClientRecord.service_center',
                                          'transform': 'clean_optional_str (обогащение карточки клиента при наличии)',
                                          'exclusion_reason': None,
                                          'header_row': 2,
                                          'nonempty_cells_below_header': 0},
                                         {'col_letter': 'D',
                                          'col_index': 4,
                                          'header': 'документов всего',
                                          'target_field': 'SliceRecord.total',
                                          'transform': 'parse_strict_int; период 2026-01-01..2026-03-31, срез '
                                                       '2026-10-20',
                                          'exclusion_reason': None,
                                          'header_row': 2,
                                          'nonempty_cells_below_header': 2928},
                                         {'col_letter': 'E',
                                          'col_index': 5,
                                          'header': 'Всего не отправлено документов',
                                          'target_field': 'SliceRecord.unsent',
                                          'transform': 'parse_strict_int; период 2026-01-01..2026-03-31, срез '
                                                       '2026-10-20',
                                          'exclusion_reason': None,
                                          'header_row': 2,
                                          'nonempty_cells_below_header': 2928}],
 'Срез 20.10.26_Дубна.xlsx::2 кв 2026': [{'col_letter': 'A',
                                          'col_index': 1,
                                          'header': 'Уникальны обезличинный идентификатор',
                                          'target_field': 'SliceRecord.source_client_id',
                                          'transform': 'Сопоставление с белым списком; неизвестные -> '
                                                       'UNKNOWN_CLIENT_ID (blocked)',
                                          'exclusion_reason': None,
                                          'header_row': 2,
                                          'nonempty_cells_below_header': 3021},
                                         {'col_letter': 'B',
                                          'col_index': 2,
                                          'header': 'код СЦ',
                                          'target_field': 'ClientRecord.service_center',
                                          'transform': 'clean_optional_str (обогащение карточки клиента при наличии)',
                                          'exclusion_reason': None,
                                          'header_row': 2,
                                          'nonempty_cells_below_header': 0},
                                         {'col_letter': 'C',
                                          'col_index': 3,
                                          'header': 'название СЦ',
                                          'target_field': 'ClientRecord.service_center',
                                          'transform': 'clean_optional_str (обогащение карточки клиента при наличии)',
                                          'exclusion_reason': None,
                                          'header_row': 2,
                                          'nonempty_cells_below_header': 0},
                                         {'col_letter': 'D',
                                          'col_index': 4,
                                          'header': 'документов всего',
                                          'target_field': 'SliceRecord.total',
                                          'transform': 'parse_strict_int; период 2026-04-01..2026-06-30, срез '
                                                       '2026-10-20',
                                          'exclusion_reason': None,
                                          'header_row': 2,
                                          'nonempty_cells_below_header': 3021},
                                         {'col_letter': 'E',
                                          'col_index': 5,
                                          'header': 'Всего не отправлено документов',
                                          'target_field': 'SliceRecord.unsent',
                                          'transform': 'parse_strict_int; период 2026-04-01..2026-06-30, срез '
                                                       '2026-10-20',
                                          'exclusion_reason': None,
                                          'header_row': 2,
                                          'nonempty_cells_below_header': 3021}],
 'Срез 20.10.26_Дубна.xlsx::3 кв 2026': [{'col_letter': 'A',
                                          'col_index': 1,
                                          'header': 'Уникальны обезличинный идентификатор',
                                          'target_field': 'SliceRecord.source_client_id',
                                          'transform': 'Сопоставление с белым списком; неизвестные -> '
                                                       'UNKNOWN_CLIENT_ID (blocked)',
                                          'exclusion_reason': None,
                                          'header_row': 2,
                                          'nonempty_cells_below_header': 3404},
                                         {'col_letter': 'B',
                                          'col_index': 2,
                                          'header': 'код СЦ',
                                          'target_field': 'ClientRecord.service_center',
                                          'transform': 'clean_optional_str (обогащение карточки клиента при наличии)',
                                          'exclusion_reason': None,
                                          'header_row': 2,
                                          'nonempty_cells_below_header': 0},
                                         {'col_letter': 'C',
                                          'col_index': 3,
                                          'header': 'название СЦ',
                                          'target_field': 'ClientRecord.service_center',
                                          'transform': 'clean_optional_str (обогащение карточки клиента при наличии)',
                                          'exclusion_reason': None,
                                          'header_row': 2,
                                          'nonempty_cells_below_header': 0},
                                         {'col_letter': 'D',
                                          'col_index': 4,
                                          'header': 'документов всего',
                                          'target_field': 'SliceRecord.total',
                                          'transform': 'parse_strict_int; период 2026-07-01..2026-09-30, срез '
                                                       '2026-10-20',
                                          'exclusion_reason': None,
                                          'header_row': 2,
                                          'nonempty_cells_below_header': 3404},
                                         {'col_letter': 'E',
                                          'col_index': 5,
                                          'header': 'Всего не отправлено документов',
                                          'target_field': 'SliceRecord.unsent',
                                          'transform': 'parse_strict_int; период 2026-07-01..2026-09-30, срез '
                                                       '2026-10-20',
                                          'exclusion_reason': None,
                                          'header_row': 2,
                                          'nonempty_cells_below_header': 3404}],
 'База должники III кв 2025 Дубна.xlsx::1-3 кв 25': [{'col_letter': 'A',
                                                      'col_index': 1,
                                                      'header': 'Уникальны обезличинный идентификатор',
                                                      'target_field': 'source_client_id',
                                                      'transform': 'str.strip(); сверка с реестром клиентов',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 295},
                                                     {'col_letter': 'B',
                                                      'col_index': 2,
                                                      'header': 'Код КЦР',
                                                      'target_field': 'ClientRecord.kcr_code',
                                                      'transform': 'clean_optional_str (обогащение карточки клиента)',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 0},
                                                     {'col_letter': 'C',
                                                      'col_index': 3,
                                                      'header': 'название СЦ',
                                                      'target_field': 'ClientRecord.service_center',
                                                      'transform': 'clean_optional_str (обогащение карточки клиента)',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 0},
                                                     {'col_letter': 'D',
                                                      'col_index': 4,
                                                      'header': 'Дата создания КЦР',
                                                      'target_field': 'ClientRecord.creation_date',
                                                      'transform': 'parse_iso_date (обогащение карточки клиента)',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 0},
                                                     {'col_letter': 'E',
                                                      'col_index': 5,
                                                      'header': 'Время от Екб',
                                                      'target_field': 'ClientRecord.time_zone_offset',
                                                      'transform': 'clean_optional_str (обогащение карточки клиента)',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 0},
                                                     {'col_letter': 'F',
                                                      'col_index': 6,
                                                      'header': 'ФИО главного админа',
                                                      'target_field': 'ClientRecord.admin_name',
                                                      'transform': 'clean_optional_str (истинное ФИО главного '
                                                                   'администратора)',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 0},
                                                     {'col_letter': 'G',
                                                      'col_index': 7,
                                                      'header': 'Телефон админа',
                                                      'target_field': 'ClientRecord.admin_phone',
                                                      'transform': 'clean_optional_str (телефон администратора из '
                                                                   'выделенной колонки)',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 0},
                                                     {'col_letter': 'H',
                                                      'col_index': 8,
                                                      'header': 'Email админа',
                                                      'target_field': 'ClientRecord.admin_email',
                                                      'transform': 'clean_optional_str (email администратора из '
                                                                   'выделенной колонки)',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 0},
                                                     {'col_letter': 'I',
                                                      'col_index': 9,
                                                      'header': '% долга за весь 2025 на 27.03.26',
                                                      'target_field': None,
                                                      'transform': 'Сохранение исходного значения и координат в '
                                                                   'source_preserved_values.jsonl при наличии значения.',
                                                      'exclusion_reason': 'Не включено в текущие '
                                                                          'ClientRecord/SliceRecord/HistoryRecord; '
                                                                          'сохранено без трактовки в '
                                                                          'source_preserved_values.jsonl. Нормализация '
                                                                          'требует согласования.',
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 295},
                                                     {'col_letter': 'J',
                                                      'col_index': 10,
                                                      'header': 'коммаентарий для претензии',
                                                      'target_field': 'HistoryRecord(claim_comment).text',
                                                      'transform': 'Дословное сохранение текста претензии (52 записи), '
                                                                   "source_column='J'",
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 52},
                                                     {'col_letter': 'K',
                                                      'col_index': 11,
                                                      'header': 'Статус на 26.05.26',
                                                      'target_field': 'HistoryRecord(historical_status).raw_status',
                                                      'transform': "occurred_at='2026-05-26', source_column='K' (184 "
                                                                   'записи)',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 184},
                                                     {'col_letter': 'L',
                                                      'col_index': 12,
                                                      'header': 'Статус на 27.03.26',
                                                      'target_field': 'HistoryRecord(historical_status).raw_status',
                                                      'transform': "occurred_at='2026-03-27', source_column='L' (295 "
                                                                   'записей)',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 295},
                                                     {'col_letter': 'M',
                                                      'col_index': 13,
                                                      'header': 'Статус на 13.03.26',
                                                      'target_field': 'HistoryRecord(historical_status).raw_status',
                                                      'transform': "occurred_at='2026-03-13', source_column='M' (295 "
                                                                   'записей)',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 295},
                                                     {'col_letter': 'N',
                                                      'col_index': 14,
                                                      'header': 'Статус на 02.03.26',
                                                      'target_field': 'HistoryRecord(historical_status).raw_status',
                                                      'transform': "occurred_at='2026-03-02', source_column='N' (32 "
                                                                   'записи)',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 32},
                                                     {'col_letter': 'O',
                                                      'col_index': 15,
                                                      'header': 'Статус на 16.01.26',
                                                      'target_field': 'HistoryRecord(historical_status).raw_status',
                                                      'transform': "occurred_at='2026-01-16', source_column='O' (294 "
                                                                   'записи)',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 294},
                                                     {'col_letter': 'P',
                                                      'col_index': 16,
                                                      'header': 'Статус на 29.12.25',
                                                      'target_field': 'HistoryRecord(historical_status).raw_status',
                                                      'transform': "occurred_at='2025-12-29', source_column='P' (295 "
                                                                   'записей)',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 295},
                                                     {'col_letter': 'Q',
                                                      'col_index': 17,
                                                      'header': 'Статус на 5.12.25',
                                                      'target_field': 'HistoryRecord(historical_status).raw_status',
                                                      'transform': "occurred_at='2025-12-05', source_column='Q' (295 "
                                                                   'записей)',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 295},
                                                     {'col_letter': 'R',
                                                      'col_index': 18,
                                                      'header': 'админ в нескольких кцр',
                                                      'target_field': None,
                                                      'transform': 'Сохранение исходного значения и координат в '
                                                                   'source_preserved_values.jsonl при наличии значения.',
                                                      'exclusion_reason': 'Не включено в текущие '
                                                                          'ClientRecord/SliceRecord/HistoryRecord; '
                                                                          'сохранено без трактовки в '
                                                                          'source_preserved_values.jsonl. Нормализация '
                                                                          'требует согласования.',
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 295},
                                                     {'col_letter': 'S',
                                                      'col_index': 19,
                                                      'header': 'Документов всего за 2025 год',
                                                      'target_field': 'SliceRecord.total',
                                                      'transform': 'parse_strict_int; базовый total для 6 '
                                                                   'ретроспективных срезов 2025 года',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 296},
                                                     {'col_letter': 'T',
                                                      'col_index': 20,
                                                      'header': 'Долг на 27.03.26',
                                                      'target_field': 'SliceRecord.unsent',
                                                      'transform': 'parse_strict_int; период 2025-01-01..2025-12-31, '
                                                                   'срез 2026-03-27',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 296},
                                                     {'col_letter': 'U',
                                                      'col_index': 21,
                                                      'header': '% Долга на 27.03.26',
                                                      'target_field': 'SliceRecord.raw_percentage',
                                                      'transform': 'parse_strict_float; срез 2026-03-27',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 296},
                                                     {'col_letter': 'V',
                                                      'col_index': 22,
                                                      'header': 'Уменьшение долга октябрь vs март',
                                                      'target_field': None,
                                                      'transform': 'Сохранение исходного значения и координат в '
                                                                   'source_preserved_values.jsonl при наличии значения.',
                                                      'exclusion_reason': 'Не включено в текущие '
                                                                          'ClientRecord/SliceRecord/HistoryRecord; '
                                                                          'сохранено без трактовки в '
                                                                          'source_preserved_values.jsonl. Нормализация '
                                                                          'требует согласования.',
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 295},
                                                     {'col_letter': 'W',
                                                      'col_index': 23,
                                                      'header': 'долг на 16.01.25',
                                                      'target_field': 'SliceRecord.unsent (rejected)',
                                                      'transform': 'Изолируется в rejected_slices.jsonl из-за '
                                                                   'противоречия даты с колонкой X (16.01.25 vs '
                                                                   '16.01.26)',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 296},
                                                     {'col_letter': 'X',
                                                      'col_index': 24,
                                                      'header': '% Долга на 16.01.26',
                                                      'target_field': 'SliceRecord.raw_percentage (rejected)',
                                                      'transform': 'Изолируется в rejected_slices.jsonl '
                                                                   '(AMBIGUOUS_AS_OF_DATE:2025-01-16_vs_2026-01-16)',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 296},
                                                     {'col_letter': 'Y',
                                                      'col_index': 25,
                                                      'header': 'Уменьшение долга ноябрь vs январь',
                                                      'target_field': None,
                                                      'transform': 'Сохранение исходного значения и координат в '
                                                                   'source_preserved_values.jsonl при наличии значения.',
                                                      'exclusion_reason': 'Не включено в текущие '
                                                                          'ClientRecord/SliceRecord/HistoryRecord; '
                                                                          'сохранено без трактовки в '
                                                                          'source_preserved_values.jsonl. Нормализация '
                                                                          'требует согласования.',
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 296},
                                                     {'col_letter': 'Z',
                                                      'col_index': 26,
                                                      'header': '% уменьшение долга ноябрь vs январь',
                                                      'target_field': None,
                                                      'transform': 'Сохранение исходного значения и координат в '
                                                                   'source_preserved_values.jsonl при наличии значения.',
                                                      'exclusion_reason': 'Не включено в текущие '
                                                                          'ClientRecord/SliceRecord/HistoryRecord; '
                                                                          'сохранено без трактовки в '
                                                                          'source_preserved_values.jsonl. Нормализация '
                                                                          'требует согласования.',
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 295},
                                                     {'col_letter': 'AA',
                                                      'col_index': 27,
                                                      'header': 'долг на 05.12.25',
                                                      'target_field': 'SliceRecord.unsent',
                                                      'transform': 'parse_strict_int; период 2025-01-01..2025-12-31, '
                                                                   'срез 2025-12-05',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 296},
                                                     {'col_letter': 'AB',
                                                      'col_index': 28,
                                                      'header': '% Долга на 05.12.25',
                                                      'target_field': 'SliceRecord.raw_percentage',
                                                      'transform': 'parse_strict_float; срез 2025-12-05',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 296},
                                                     {'col_letter': 'AC',
                                                      'col_index': 29,
                                                      'header': 'Уменьшение долга ноябрь vs декабрь',
                                                      'target_field': None,
                                                      'transform': 'Сохранение исходного значения и координат в '
                                                                   'source_preserved_values.jsonl при наличии значения.',
                                                      'exclusion_reason': 'Не включено в текущие '
                                                                          'ClientRecord/SliceRecord/HistoryRecord; '
                                                                          'сохранено без трактовки в '
                                                                          'source_preserved_values.jsonl. Нормализация '
                                                                          'требует согласования.',
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 296},
                                                     {'col_letter': 'AD',
                                                      'col_index': 30,
                                                      'header': '% уменьшение долга ноябрь vs декабрь',
                                                      'target_field': None,
                                                      'transform': 'Сохранение исходного значения и координат в '
                                                                   'source_preserved_values.jsonl при наличии значения.',
                                                      'exclusion_reason': 'Не включено в текущие '
                                                                          'ClientRecord/SliceRecord/HistoryRecord; '
                                                                          'сохранено без трактовки в '
                                                                          'source_preserved_values.jsonl. Нормализация '
                                                                          'требует согласования.',
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 296},
                                                     {'col_letter': 'AE',
                                                      'col_index': 31,
                                                      'header': 'долг на 05.11.25',
                                                      'target_field': 'SliceRecord.unsent',
                                                      'transform': 'parse_strict_int; период 2025-01-01..2025-12-31, '
                                                                   'срез 2025-11-05',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 296},
                                                     {'col_letter': 'AF',
                                                      'col_index': 32,
                                                      'header': '% Долга на 05.11.25',
                                                      'target_field': 'SliceRecord.raw_percentage',
                                                      'transform': 'parse_strict_float; срез 2025-11-05',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 296},
                                                     {'col_letter': 'AG',
                                                      'col_index': 33,
                                                      'header': 'долг на 31.10.25',
                                                      'target_field': 'SliceRecord.unsent',
                                                      'transform': 'parse_strict_int; период 2025-01-01..2025-12-31, '
                                                                   'срез 2025-10-31 (#N/A -> rejected)',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 295},
                                                     {'col_letter': 'AH',
                                                      'col_index': 34,
                                                      'header': '% Долга на 31.10.25',
                                                      'target_field': 'SliceRecord.raw_percentage',
                                                      'transform': 'parse_strict_float; срез 2025-10-31 (#N/A -> '
                                                                   'rejected)',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 295},
                                                     {'col_letter': 'AI',
                                                      'col_index': 35,
                                                      'header': 'долг на 20.10.25',
                                                      'target_field': 'SliceRecord.unsent',
                                                      'transform': 'parse_strict_int; период 2025-01-01..2025-12-31, '
                                                                   'срез 2025-10-20',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 295},
                                                     {'col_letter': 'AJ',
                                                      'col_index': 36,
                                                      'header': 'отклонено',
                                                      'target_field': None,
                                                      'transform': 'Сохранение исходного значения и координат в '
                                                                   'source_preserved_values.jsonl при наличии значения.',
                                                      'exclusion_reason': 'Не включено в текущие '
                                                                          'ClientRecord/SliceRecord/HistoryRecord; '
                                                                          'сохранено без трактовки в '
                                                                          'source_preserved_values.jsonl. Нормализация '
                                                                          'требует согласования.',
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 295},
                                                     {'col_letter': 'AK',
                                                      'col_index': 37,
                                                      'header': 'на исправлении',
                                                      'target_field': None,
                                                      'transform': 'Сохранение исходного значения и координат в '
                                                                   'source_preserved_values.jsonl при наличии значения.',
                                                      'exclusion_reason': 'Не включено в текущие '
                                                                          'ClientRecord/SliceRecord/HistoryRecord; '
                                                                          'сохранено без трактовки в '
                                                                          'source_preserved_values.jsonl. Нормализация '
                                                                          'требует согласования.',
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 295},
                                                     {'col_letter': 'AL',
                                                      'col_index': 38,
                                                      'header': '% долга',
                                                      'target_field': None,
                                                      'transform': 'Сохранение исходного значения и координат в '
                                                                   'source_preserved_values.jsonl при наличии значения.',
                                                      'exclusion_reason': 'Не включено в текущие '
                                                                          'ClientRecord/SliceRecord/HistoryRecord; '
                                                                          'сохранено без трактовки в '
                                                                          'source_preserved_values.jsonl. Нормализация '
                                                                          'требует согласования.',
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 295},
                                                     {'col_letter': 'AM',
                                                      'col_index': 39,
                                                      'header': 'отправлено',
                                                      'target_field': None,
                                                      'transform': 'Сохранение исходного значения и координат в '
                                                                   'source_preserved_values.jsonl при наличии значения.',
                                                      'exclusion_reason': 'Не включено в текущие '
                                                                          'ClientRecord/SliceRecord/HistoryRecord; '
                                                                          'сохранено без трактовки в '
                                                                          'source_preserved_values.jsonl. Нормализация '
                                                                          'требует согласования.',
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 295},
                                                     {'col_letter': 'AN',
                                                      'col_index': 40,
                                                      'header': 'принято',
                                                      'target_field': None,
                                                      'transform': 'Сохранение исходного значения и координат в '
                                                                   'source_preserved_values.jsonl при наличии значения.',
                                                      'exclusion_reason': 'Не включено в текущие '
                                                                          'ClientRecord/SliceRecord/HistoryRecord; '
                                                                          'сохранено без трактовки в '
                                                                          'source_preserved_values.jsonl. Нормализация '
                                                                          'требует согласования.',
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 295},
                                                     {'col_letter': 'AO',
                                                      'col_index': 41,
                                                      'header': 'Статус на 20.10.25',
                                                      'target_field': 'HistoryRecord(historical_status).raw_status',
                                                      'transform': "occurred_at='2025-10-20', source_column='AO' (295 "
                                                                   'записей)',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 295},
                                                     {'col_letter': 'AP',
                                                      'col_index': 42,
                                                      'header': 'Статус на 05.11.25',
                                                      'target_field': 'HistoryRecord(historical_status).raw_status',
                                                      'transform': "occurred_at='2025-11-05', source_column='AP' (295 "
                                                                   'записей)',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 295},
                                                     {'col_letter': 'AQ',
                                                      'col_index': 43,
                                                      'header': 'Арина 14.10.25',
                                                      'target_field': None,
                                                      'transform': 'Сохранение исходного значения и координат в '
                                                                   'source_preserved_values.jsonl при наличии значения.',
                                                      'exclusion_reason': 'Не включено в текущие '
                                                                          'ClientRecord/SliceRecord/HistoryRecord; '
                                                                          'сохранено без трактовки в '
                                                                          'source_preserved_values.jsonl. Нормализация '
                                                                          'требует согласования.',
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 295},
                                                     {'col_letter': 'AR',
                                                      'col_index': 44,
                                                      'header': 'Отправляли посылку без описи',
                                                      'target_field': None,
                                                      'transform': 'Сохранение исходного значения и координат в '
                                                                   'source_preserved_values.jsonl при наличии значения.',
                                                      'exclusion_reason': 'Не включено в текущие '
                                                                          'ClientRecord/SliceRecord/HistoryRecord; '
                                                                          'сохранено без трактовки в '
                                                                          'source_preserved_values.jsonl. Нормализация '
                                                                          'требует согласования.',
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 295},
                                                     {'col_letter': 'AS',
                                                      'col_index': 45,
                                                      'header': 'Ответственный ООВС',
                                                      'target_field': None,
                                                      'transform': 'Сохранение исходного значения и координат в '
                                                                   'source_preserved_values.jsonl при наличии значения.',
                                                      'exclusion_reason': 'Нет непустых значений ниже заголовка в '
                                                                          'предоставленном файле.',
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 0},
                                                     {'col_letter': 'AT',
                                                      'col_index': 46,
                                                      'header': 'Комментарии ООВС',
                                                      'target_field': 'HistoryRecord(first_contact).text',
                                                      'transform': 'Дословный текст комментария первого касания '
                                                                   "(source_column='AT-AX')",
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 295},
                                                     {'col_letter': 'AU',
                                                      'col_index': 47,
                                                      'header': 'Статус звонка',
                                                      'target_field': 'HistoryRecord(first_contact).raw_status',
                                                      'transform': 'Дословный статус первого звонка '
                                                                   "(source_column='AT-AX')",
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 295},
                                                     {'col_letter': 'AV',
                                                      'col_index': 48,
                                                      'header': 'Дата перезвона',
                                                      'target_field': 'HistoryRecord(first_contact).next_contact_date '
                                                                      '/ raw_next_contact',
                                                      'transform': 'parse_iso_date + сохранение некалендарного текста '
                                                                   'и NON_DATE_NEXT_CONTACT',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 0},
                                                     {'col_letter': 'AW',
                                                      'col_index': 49,
                                                      'header': 'Срок отправки от клиента',
                                                      'target_field': 'HistoryRecord(first_contact).raw_send_deadline',
                                                      'transform': 'Сохраняется в raw_send_deadline и в объединённом '
                                                                   'raw_deadline',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 282},
                                                     {'col_letter': 'AX',
                                                      'col_index': 50,
                                                      'header': 'Дата от клиента',
                                                      'target_field': 'HistoryRecord(first_contact).raw_client_date',
                                                      'transform': 'Сохраняется в raw_client_date (62 записи) и в '
                                                                   'объединённом raw_deadline',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 62},
                                                     {'col_letter': 'AY',
                                                      'col_index': 51,
                                                      'header': 'Звоним в январе?',
                                                      'target_field': None,
                                                      'transform': 'Сохранение исходного значения и координат в '
                                                                   'source_preserved_values.jsonl при наличии значения.',
                                                      'exclusion_reason': 'Не включено в текущие '
                                                                          'ClientRecord/SliceRecord/HistoryRecord; '
                                                                          'сохранено без трактовки в '
                                                                          'source_preserved_values.jsonl. Нормализация '
                                                                          'требует согласования.',
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 295},
                                                     {'col_letter': 'AZ',
                                                      'col_index': 52,
                                                      'header': 'Ответственный ООВС',
                                                      'target_field': None,
                                                      'transform': 'Сохранение исходного значения и координат в '
                                                                   'source_preserved_values.jsonl при наличии значения.',
                                                      'exclusion_reason': 'Нет непустых значений ниже заголовка в '
                                                                          'предоставленном файле.',
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 0},
                                                     {'col_letter': 'BA',
                                                      'col_index': 53,
                                                      'header': 'Комментарий к повторному звонку',
                                                      'target_field': 'HistoryRecord(second_contact).text',
                                                      'transform': 'Дословный комментарий повторного звонка в январе '
                                                                   "(source_column='BA-BE')",
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 152},
                                                     {'col_letter': 'BB',
                                                      'col_index': 54,
                                                      'header': 'Статус звонка в январе',
                                                      'target_field': 'HistoryRecord(second_contact).raw_status',
                                                      'transform': 'Дословный статус повторного звонка в январе (278 '
                                                                   'записей)',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 278},
                                                     {'col_letter': 'BC',
                                                      'col_index': 55,
                                                      'header': 'Дата перезвона',
                                                      'target_field': 'HistoryRecord(second_contact).next_contact_date '
                                                                      '/ raw_next_contact',
                                                      'transform': 'parse_iso_date + сохранение сырого текста и '
                                                                   'NON_DATE_NEXT_CONTACT',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 5},
                                                     {'col_letter': 'BD',
                                                      'col_index': 56,
                                                      'header': 'Срок отправки по клиенту',
                                                      'target_field': 'HistoryRecord(second_contact).raw_send_deadline',
                                                      'transform': 'Сохраняется в raw_send_deadline (229 записей) и в '
                                                                   'raw_deadline',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 229},
                                                     {'col_letter': 'BE',
                                                      'col_index': 57,
                                                      'header': 'Срок от клиента',
                                                      'target_field': 'HistoryRecord(second_contact).raw_client_date',
                                                      'transform': 'Сохраняется в raw_client_date (19 записей) и в '
                                                                   'raw_deadline',
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 19},
                                                     {'col_letter': 'BF',
                                                      'col_index': 58,
                                                      'header': 'Инфа по документам от клиента',
                                                      'target_field': 'HistoryRecord(incident_or_expert_request).raw_deadline '
                                                                      '/ text',
                                                      'transform': 'Сохраняется в истории инцидентов/экспертных '
                                                                   "запросов (source_column='BF-BH')",
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 1},
                                                     {'col_letter': 'BG',
                                                      'col_index': 59,
                                                      'header': '№ Инцидента по отсрочке',
                                                      'target_field': 'HistoryRecord(incident_or_expert_request).text',
                                                      'transform': 'Сохраняется в истории инцидентов/экспертных '
                                                                   "запросов (source_column='BF-BH')",
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 1},
                                                     {'col_letter': 'BH',
                                                      'col_index': 60,
                                                      'header': 'Запрос эксперта',
                                                      'target_field': 'HistoryRecord(incident_or_expert_request).raw_status '
                                                                      '/ text',
                                                      'transform': 'Сохраняется в истории инцидентов/экспертных '
                                                                   "запросов (source_column='BF-BH')",
                                                      'exclusion_reason': None,
                                                      'header_row': 1,
                                                      'nonempty_cells_below_header': 6}],
 'База I кв 2026 Дубна.xlsx::Все КЦР': [{'col_letter': 'A',
                                         'col_index': 1,
                                         'header': 'Уникальны обезличинный идентификатор',
                                         'target_field': 'source_client_id',
                                         'transform': 'str.strip(); сверка с реестром клиентов',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 115},
                                        {'col_letter': 'B',
                                         'col_index': 2,
                                         'header': 'Код КЦР',
                                         'target_field': 'ClientRecord.kcr_code',
                                         'transform': 'clean_optional_str (обогащение карточки клиента)',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 0},
                                        {'col_letter': 'C',
                                         'col_index': 3,
                                         'header': None,
                                         'target_field': None,
                                         'transform': 'Сохранение исходного значения и координат в '
                                                      'source_preserved_values.jsonl при наличии значения.',
                                         'exclusion_reason': 'Нет непустых значений ниже заголовка в предоставленном '
                                                             'файле.',
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 0},
                                        {'col_letter': 'D',
                                         'col_index': 4,
                                         'header': 'Название',
                                         'target_field': 'ClientRecord.organization_name',
                                         'transform': 'clean_optional_str (обогащение карточки клиента)',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 0},
                                        {'col_letter': 'E',
                                         'col_index': 5,
                                         'header': 'СЦ',
                                         'target_field': 'ClientRecord.service_center',
                                         'transform': 'clean_optional_str (обогащение карточки клиента)',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 0},
                                        {'col_letter': 'F',
                                         'col_index': 6,
                                         'header': None,
                                         'target_field': None,
                                         'transform': 'Сохранение исходного значения и координат в '
                                                      'source_preserved_values.jsonl при наличии значения.',
                                         'exclusion_reason': 'Нет непустых значений ниже заголовка в предоставленном '
                                                             'файле.',
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 0},
                                        {'col_letter': 'G',
                                         'col_index': 7,
                                         'header': None,
                                         'target_field': None,
                                         'transform': 'Сохранение исходного значения и координат в '
                                                      'source_preserved_values.jsonl при наличии значения.',
                                         'exclusion_reason': 'Нет непустых значений ниже заголовка в предоставленном '
                                                             'файле.',
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 0},
                                        {'col_letter': 'H',
                                         'col_index': 8,
                                         'header': None,
                                         'target_field': None,
                                         'transform': 'Сохранение исходного значения и координат в '
                                                      'source_preserved_values.jsonl при наличии значения.',
                                         'exclusion_reason': 'Нет непустых значений ниже заголовка в предоставленном '
                                                             'файле.',
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 0},
                                        {'col_letter': 'I',
                                         'col_index': 9,
                                         'header': None,
                                         'target_field': None,
                                         'transform': 'Сохранение исходного значения и координат в '
                                                      'source_preserved_values.jsonl при наличии значения.',
                                         'exclusion_reason': 'Нет непустых значений ниже заголовка в предоставленном '
                                                             'файле.',
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 0},
                                        {'col_letter': 'J',
                                         'col_index': 10,
                                         'header': 'Время от Екб',
                                         'target_field': 'ClientRecord.time_zone_offset',
                                         'transform': 'clean_optional_str (обогащение 115 карточек клиентов с '
                                                      'фиксацией attribute_provenance)',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 115},
                                        {'col_letter': 'K',
                                         'col_index': 11,
                                         'header': 'ФИО главного админа',
                                         'target_field': 'ClientRecord.admin_name',
                                         'transform': 'clean_optional_str (ФИО главного администратора, отдельно от '
                                                      'manager_name)',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 0},
                                        {'col_letter': 'L',
                                         'col_index': 12,
                                         'header': 'Телефон админа',
                                         'target_field': 'ClientRecord.admin_phone',
                                         'transform': 'clean_optional_str (телефон администратора из выделенного '
                                                      'столбца)',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 0},
                                        {'col_letter': 'M',
                                         'col_index': 13,
                                         'header': 'Email админа',
                                         'target_field': 'ClientRecord.admin_email',
                                         'transform': 'clean_optional_str (email администратора из выделенного '
                                                      'столбца)',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 0},
                                        {'col_letter': 'N',
                                         'col_index': 14,
                                         'header': 'Спич',
                                         'target_field': None,
                                         'transform': 'Сохранение исходного значения и координат в '
                                                      'source_preserved_values.jsonl при наличии значения.',
                                         'exclusion_reason': 'Не включено в текущие '
                                                             'ClientRecord/SliceRecord/HistoryRecord; сохранено без '
                                                             'трактовки в source_preserved_values.jsonl. Нормализация '
                                                             'требует согласования.',
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 115},
                                        {'col_letter': 'O',
                                         'col_index': 15,
                                         'header': 'стаус на 21.05.26',
                                         'target_field': 'HistoryRecord(historical_status).raw_status',
                                         'transform': "occurred_at='2026-05-21', source_column='O' (115 записей)",
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 115},
                                        {'col_letter': 'P',
                                         'col_index': 16,
                                         'header': 'ИНН',
                                         'target_field': 'ClientRecord.inn',
                                         'transform': 'clean_optional_str (обогащение карточки клиента)',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 0},
                                        {'col_letter': 'Q',
                                         'col_index': 17,
                                         'header': 'КПП',
                                         'target_field': 'ClientRecord.kpp',
                                         'transform': 'clean_optional_str (обогащение карточки клиента)',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 0},
                                        {'col_letter': 'R',
                                         'col_index': 18,
                                         'header': 'Действует ли КЦР',
                                         'target_field': 'ClientRecord.is_active',
                                         'transform': 'clean_optional_str (обогащение карточки клиента)',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 115},
                                        {'col_letter': 'S',
                                         'col_index': 19,
                                         'header': 'Блокировка в сервисе',
                                         'target_field': 'ClientRecord.is_blocked',
                                         'transform': 'clean_optional_str (обогащение карточки клиента)',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 115},
                                        {'col_letter': 'T',
                                         'col_index': 20,
                                         'header': 'ФИО менеджера',
                                         'target_field': 'ClientRecord.manager_name',
                                         'transform': 'clean_optional_str (сохраняется строго в manager_name)',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 0},
                                        {'col_letter': 'U',
                                         'col_index': 21,
                                         'header': 'ФИО менеджера из КП',
                                         'target_field': 'ClientRecord.manager_name',
                                         'transform': 'Резервное поле manager_name, если T пусто',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 0},
                                        {'col_letter': 'V',
                                         'col_index': 22,
                                         'header': 'Всего документов весь 2025',
                                         'target_field': 'SliceRecord.total',
                                         'transform': 'parse_strict_int; период 2025-01-01..2025-12-31, срез '
                                                      '2026-05-21',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 115},
                                        {'col_letter': 'W',
                                         'col_index': 23,
                                         'header': 'Долг весь 2025 на 21.05.26',
                                         'target_field': 'SliceRecord.unsent',
                                         'transform': 'parse_strict_int; период 2025-01-01..2025-12-31, срез '
                                                      '2026-05-21',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 115},
                                        {'col_letter': 'X',
                                         'col_index': 24,
                                         'header': '% Долга весь 2025 на 21.05.26',
                                         'target_field': 'SliceRecord.raw_percentage',
                                         'transform': 'parse_strict_float; срез 2026-05-21',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 115},
                                        {'col_letter': 'Y',
                                         'col_index': 25,
                                         'header': 'Всего документов 2-4 кв 2025',
                                         'target_field': 'SliceRecord.total',
                                         'transform': 'parse_strict_int; период 2025-04-01..2025-12-31, срез '
                                                      '2026-05-21',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 114},
                                        {'col_letter': 'Z',
                                         'col_index': 26,
                                         'header': 'Долг2-4 кв 2025 на 21.05.26',
                                         'target_field': 'SliceRecord.unsent',
                                         'transform': 'parse_strict_int; период 2025-04-01..2025-12-31, срез '
                                                      '2026-05-21',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 114},
                                        {'col_letter': 'AA',
                                         'col_index': 27,
                                         'header': '% Долга 2-4кв 2025 на 21.05.26',
                                         'target_field': 'SliceRecord.raw_percentage',
                                         'transform': 'parse_strict_float; срез 2026-05-21',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 114},
                                        {'col_letter': 'AB',
                                         'col_index': 28,
                                         'header': 'Всего документов 1 кв 2026',
                                         'target_field': 'SliceRecord.total',
                                         'transform': 'parse_strict_int; период 2026-01-01..2026-03-31, срез '
                                                      '2026-05-21',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 107},
                                        {'col_letter': 'AC',
                                         'col_index': 29,
                                         'header': 'Долг 1 кв 2026 на 21.05.26',
                                         'target_field': 'SliceRecord.unsent',
                                         'transform': 'parse_strict_int; период 2026-01-01..2026-03-31, срез '
                                                      '2026-05-21',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 107},
                                        {'col_letter': 'AD',
                                         'col_index': 30,
                                         'header': '% Долга 1 кв 2026 на 21.05.26',
                                         'target_field': 'SliceRecord.raw_percentage',
                                         'transform': 'parse_strict_float; срез 2026-05-21',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 107},
                                        {'col_letter': 'AE',
                                         'col_index': 31,
                                         'header': 'Всего документов 2 кв 2025-1 кв 2026 на 21.05.26',
                                         'target_field': 'SliceRecord.total',
                                         'transform': 'parse_strict_int; период 2025-04-01..2026-03-31, срез '
                                                      '2026-05-21',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 115},
                                        {'col_letter': 'AF',
                                         'col_index': 32,
                                         'header': 'Долг 2 кв 2025-1 кв 2026 на 21.05.26',
                                         'target_field': 'SliceRecord.unsent',
                                         'transform': 'parse_strict_int; период 2025-04-01..2026-03-31, срез '
                                                      '2026-05-21',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 115},
                                        {'col_letter': 'AG',
                                         'col_index': 33,
                                         'header': '% Долга 2 кв 2025-1 кв 2026 на 21.05.26',
                                         'target_field': 'SliceRecord.raw_percentage',
                                         'transform': 'parse_strict_float; срез 2026-05-21',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 113},
                                        {'col_letter': 'AH',
                                         'col_index': 34,
                                         'header': 'Ответственный ООВС',
                                         'target_field': None,
                                         'transform': 'Сохранение исходного значения и координат в '
                                                      'source_preserved_values.jsonl при наличии значения.',
                                         'exclusion_reason': 'Нет непустых значений ниже заголовка в предоставленном '
                                                             'файле.',
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 0},
                                        {'col_letter': 'AI',
                                         'col_index': 35,
                                         'header': 'Комментарии ООВС',
                                         'target_field': 'HistoryRecord(operator_contact).text',
                                         'transform': 'Дословный комментарий оператора I кв. 2026 (115 записей, '
                                                      "source_column='AI-AM')",
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 115},
                                        {'col_letter': 'AJ',
                                         'col_index': 36,
                                         'header': 'Статус звонка',
                                         'target_field': 'HistoryRecord(operator_contact).raw_status',
                                         'transform': 'Дословный статус звонка I кв. 2026 (115 записей, '
                                                      "source_column='AI-AM')",
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 115},
                                        {'col_letter': 'AK',
                                         'col_index': 37,
                                         'header': 'Дата перезвона',
                                         'target_field': 'HistoryRecord(operator_contact).next_contact_date / '
                                                         'raw_next_contact',
                                         'transform': 'parse_iso_date + сохранение некалендарного текста и '
                                                      'NON_DATE_NEXT_CONTACT',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 0},
                                        {'col_letter': 'AL',
                                         'col_index': 38,
                                         'header': 'Срок отправки от клиента',
                                         'target_field': 'HistoryRecord(operator_contact).raw_send_deadline',
                                         'transform': 'Сохраняется в raw_send_deadline (93 записи) и в raw_deadline',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 93},
                                        {'col_letter': 'AM',
                                         'col_index': 39,
                                         'header': 'Дата от клиента',
                                         'target_field': 'HistoryRecord(operator_contact).raw_client_date',
                                         'transform': 'Сохраняется в raw_client_date и в raw_deadline',
                                         'exclusion_reason': None,
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 0},
                                        {'col_letter': 'AN',
                                         'col_index': 40,
                                         'header': '№ Инцидента по отсрочке',
                                         'target_field': None,
                                         'transform': 'Сохранение исходного значения и координат в '
                                                      'source_preserved_values.jsonl при наличии значения.',
                                         'exclusion_reason': 'Нет непустых значений ниже заголовка в предоставленном '
                                                             'файле.',
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 0},
                                        {'col_letter': 'AO',
                                         'col_index': 41,
                                         'header': 'Запрос эксперта',
                                         'target_field': None,
                                         'transform': 'Сохранение исходного значения и координат в '
                                                      'source_preserved_values.jsonl при наличии значения.',
                                         'exclusion_reason': 'Нет непустых значений ниже заголовка в предоставленном '
                                                             'файле.',
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 0},
                                        {'col_letter': 'AP',
                                         'col_index': 42,
                                         'header': None,
                                         'target_field': None,
                                         'transform': 'Сохранение исходного значения и координат в '
                                                      'source_preserved_values.jsonl при наличии значения.',
                                         'exclusion_reason': 'Нет непустых значений ниже заголовка в предоставленном '
                                                             'файле.',
                                         'header_row': 2,
                                         'nonempty_cells_below_header': 0}]}


# =====================================================================
# Вспомогательные функции валидации и преобразования
# =====================================================================

def clean_optional_str(val: Any) -> Optional[str]:
    if val is None:
        return None
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, float) and math.isfinite(val) and val.is_integer():
        return str(int(val))
    s = str(val).strip()
    return s if s != "" else None


def parse_iso_date(val: Any) -> Optional[str]:
    if val is None:
        return None
    if isinstance(val, (datetime, date)):
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
    if val_str.startswith("="):
        return None, f"FORMULA_WITHOUT_CACHED_VALUE:{val}"
    try:
        return int(val_str), None
    except ValueError:
        pass
    try:
        d_val = Decimal(val_str)
        if not d_val.is_finite():
            return None, f"NON_FINITE_NUMBER:{val}"
        if d_val == int(d_val):
            return int(d_val), None
        return None, f"FRACTIONAL_NUMBER:{val}"
    except (InvalidOperation, ValueError, OverflowError):
        return None, f"INVALID_NUMERIC_STRING:{val}"


def parse_strict_float(val: Any) -> Tuple[Optional[float], Optional[str]]:
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
    if val_str.startswith("="):
        return None, f"FORMULA_WITHOUT_CACHED_VALUE:{val}"
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
    """Нормализует тире/дефисы, римские цифры кварталов, границы букв/цифр и пробелы."""
    s = (text or "").strip().lower()
    s = re.sub(r"\bквартал[а-я]*\b", "кв", s)
    s = re.sub(r"\bкв\.", "кв", s)
    s = re.sub(r"\biv\s*кв\b", "4 кв", s)
    s = re.sub(r"\biii\s*кв\b", "3 кв", s)
    s = re.sub(r"\bii\s*кв\b", "2 кв", s)
    s = re.sub(r"\bi\s*кв\b", "1 кв", s)
    s = re.sub(r"([а-яa-z])(\d)", r"\1 \2", s)
    s = re.sub(r"(\d)([а-яa-z])", r"\1 \2", s)
    s = re.sub(r"\s*[\-\u2010-\u2015\u2212]+\s*", "-", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def extract_period_from_header(raw_header: str) -> str:
    """
    Извлекает полное нормализованное выражение отчётного периода из заголовка колонки
    (включая начало и окончание интервала), удаляя названия метрик ('Долг', 'Всего документов', '%')
    и дату среза ('на ДД.ММ.ГГ').
    """
    s = (raw_header or "").strip().lower()
    # Удаляем 'на ДД.ММ.ГГ(ГГ)' и отдельные даты среза
    s = re.sub(r"\bна\s+\d{1,2}\.\d{2}\.\d{2,4}\b", " ", s)
    s = re.sub(r"\b\d{1,2}\.\d{2}\.\d{2,4}\b", " ", s)
    # Нормализуем обозначения кварталов и римские цифры
    s = re.sub(r"\bквартал[а-я]*\b", "кв", s)
    s = re.sub(r"\bкв\.", "кв", s)
    s = re.sub(r"([а-яa-z])(\d)", r"\1 \2", s)
    s = re.sub(r"(\d)([а-яa-z])", r"\1 \2", s)
    s = re.sub(r"\biv\s*кв\b", "4 кв", s)
    s = re.sub(r"\biii\s*кв\b", "3 кв", s)
    s = re.sub(r"\bii\s*кв\b", "2 кв", s)
    s = re.sub(r"\bi\s*кв\b", "1 кв", s)
    # Удаляем служебные слова метрик и слово 'год'/'г.'
    s = re.sub(r"%+", " ", s)
    s = re.sub(r"\b(всего|документов|документам|документы|долга|долг|по|за|на|год[ауе]?|г\.?)\b", " ", s)
    # Приводим разделители интервала (тире, дефисы, слэши, 'по') к единому дефису '-'
    s = re.sub(r"\s*(?:[\-\u2010-\u2015\u2212/]+|\bпо\b)\s*", "-", s)
    s = re.sub(r"\s+", " ", s).strip(" -")
    # Канонизация '2025' -> 'весь 2025' и '2-4' -> '2-4 кв 2025' для единообразного сравнения полного периода
    if s == "2025":
        s = "весь 2025"
    elif s == "2-4" or s == "2-4 кв":
        s = "2-4 кв 2025"
    return s


def verify_header_exact_period(raw_header: str, period_token: str) -> bool:
    """
    Сравнивает полный извлечённый из заголовка период (включая его начало и окончание)
    с ожидаемым периодом колонки. Исключает принятие 'весь 2025 — II квартал 2026'
    или 'весь 2025 год - 2 кв 2026' как 'весь 2025'.
    """
    extracted_h = extract_period_from_header(raw_header)
    extracted_p = extract_period_from_header(period_token)
    if not extracted_h or not extracted_p:
        return False
    return extracted_h == extracted_p


def verify_header_tokens(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    h_row: int,
    expected_tokens: Dict[int, List[str]],
    filename: str,
    sheetname: str
) -> None:
    for col_idx, tokens in expected_tokens.items():
        raw_val = str(ws.cell(h_row, col_idx).value or "")
        norm_h = normalize_header_for_match(raw_val)
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        for tok in tokens:
            norm_tok = normalize_header_for_match(tok)
            if any(kw in norm_tok for kw in ("2025", "2026", "2-4", "1 кв", "2 кв", "3-4 кв", "3 кв", "4 кв")):
                if not verify_header_exact_period(raw_val, tok):
                    raise ValueError(
                        f"Структура {filename} (лист '{sheetname}') нарушена в колонке {col_letter} (столбец {col_idx}): "
                        f"ожидался период '{tok}', фактически '{ws.cell(h_row, col_idx).value}'"
                    )
            else:
                if norm_tok not in norm_h:
                    raise ValueError(
                        f"Структура {filename} (лист '{sheetname}') нарушена в колонке {col_letter} (столбец {col_idx}): "
                        f"ожидался фрагмент '{tok}', фактически '{ws.cell(h_row, col_idx).value}'"
                    )


def extract_dates_from_header(header_text: str) -> List[str]:
    matches = re.findall(r"\b(\d{1,2}\.\d{2}\.\d{2,4})\b", header_text or "")
    res = []
    for m in matches:
        iso = parse_iso_date(m)
        if iso:
            res.append(iso)
    return res


def get_cell_value_with_formula_check(
    ws_val: openpyxl.worksheet.worksheet.Worksheet,
    ws_f: Optional[openpyxl.worksheet.worksheet.Worksheet],
    r: int,
    c: int
) -> Any:
    val = ws_val.cell(r, c).value
    if val is None and ws_f is not None:
        f_val = ws_f.cell(r, c).value
        if isinstance(f_val, str) and f_val.strip().startswith("="):
            return f_val.strip()
    return val


def set_client_attr(
    c_rec: ClientRecord,
    attr_name: str,
    val: Optional[str],
    filename: str,
    sheetname: str,
    row: int,
    col_letter: str,
    overwrite: bool = False
) -> None:
    if val is not None and (overwrite or getattr(c_rec, attr_name) is None):
        setattr(c_rec, attr_name, val)
        c_rec.attribute_provenance[attr_name] = Provenance(
            source_file=filename,
            source_sheet=sheetname,
            source_row=row,
            source_column=col_letter
        )


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
    base_rejection: Optional[str] = None,
    ws_f: Optional[openpyxl.worksheet.worksheet.Worksheet] = None
) -> Optional[SliceRecord]:
    raw_tot = get_cell_value_with_formula_check(ws, ws_f, r, c_tot)
    raw_uns = get_cell_value_with_formula_check(ws, ws_f, r, c_uns)
    raw_pct = get_cell_value_with_formula_check(ws, ws_f, r, c_pct) if c_pct else None

    tot_empty = raw_tot is None or (isinstance(raw_tot, str) and raw_tot.strip() == "")
    uns_empty = raw_uns is None or (isinstance(raw_uns, str) and raw_uns.strip() == "")

    if tot_empty and uns_empty:
        if raw_pct is not None and str(raw_pct).strip() != "":
            issues.append(QualityIssue(
                code="ORPHAN_PERCENTAGE_WITHOUT_COUNTS",
                file=filename, sheet=sheetname, row=r, col=openpyxl.utils.get_column_letter(c_pct),
                raw_value=str(raw_pct),
                description="Указан процент при отсутствующих значениях total и unsent",
                action="requires_review"
            ))
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
    SLICE_SPECS = [
        (15, 16, 17, "2025-01-01", "2025-12-31", "2026-09-29", "весь 2025", "29.09.26"),
        (15, 18, 19, "2025-01-01", "2025-12-31", "2026-07-21", "весь 2025", "21.07.26"),
        (15, 20, 21, "2025-01-01", "2025-12-31", "2026-07-07", "весь 2025", "07.07.26"),
        (15, 22, 23, "2025-01-01", "2025-12-31", "2026-05-21", "весь 2025", "21.05.26"),
        (24, 25, 26, "2025-04-01", "2025-12-31", "2026-07-21", "2-4 кв 2025", "21.07.26"),
        (24, 27, 28, "2025-04-01", "2025-12-31", "2026-07-07", "2-4 кв 2025", "07.07.26"),
        (24, 29, 30, "2025-04-01", "2025-12-31", "2026-05-21", "2-4 кв 2025", "21.05.26"),
        (31, 32, 33, "2026-01-01", "2026-03-31", "2026-09-29", "1 кв 2026", "29.09.26"),
        (31, 34, 35, "2026-01-01", "2026-03-31", "2026-07-21", "1 кв 2026", "21.07.26"),
        (31, 36, 37, "2026-01-01", "2026-03-31", "2026-07-07", "1 кв 2026", "07.07.26"),
        (31, 38, 39, "2026-01-01", "2026-03-31", "2026-05-21", "1 кв 2026", "21.05.26"),
        (40, 41, 42, "2025-04-01", "2026-03-31", "2026-07-21", "2 кв 2025-1 кв 2026", "21.07.26"),
        (40, 43, 44, "2025-04-01", "2026-03-31", "2026-07-07", "2 кв 2025-1 кв 2026", "07.07.26"),
        (40, 45, 46, "2025-04-01", "2026-03-31", "2026-05-21", "2 кв 2025-1 кв 2026", "21.05.26"),
        (47, 48, 49, "2025-07-01", "2025-12-31", "2026-09-29", "3-4 кв 2025", "29.09.26"),
        (47, 50, 51, "2025-07-01", "2025-12-31", "2026-07-21", "3-4 кв 2025", "21.07.26"),
        (52, 53, 54, "2026-04-01", "2026-06-30", "2026-09-29", "2 кв 2026", "29.09.26"),
        (52, 55, 56, "2026-04-01", "2026-06-30", "2026-07-21", "2 кв 2026", "21.07.26"),
        (57, 58, 59, "2025-07-01", "2026-06-30", "2026-09-29", "3 кв 2025-2 кв 2026", "29.09.26"),
        (57, 60, 61, "2025-07-01", "2026-06-30", "2026-07-21", "3 кв 2025-2 кв 2026", "21.07.26"),
        (62, 63, 64, "2025-01-01", "2026-06-30", "2026-09-29", "весь 2025-2 кв 2026", "29.09.26"),
        (62, 65, 66, "2025-01-01", "2026-06-30", "2026-07-21", "весь 2025-2 кв 2026", "21.07.26"),
    ]

    @classmethod
    def parse(cls, filepath: str) -> Tuple[Dict[str, ClientRecord], List[SliceRecord], List[HistoryRecord], List[QualityIssue]]:
        filename = os.path.basename(filepath)
        wb = openpyxl.load_workbook(filepath, data_only=True)
        wb_f = openpyxl.load_workbook(filepath, data_only=False)
        try:
            if "Все КЦР" not in wb.sheetnames:
                raise ValueError(f"В файле {filename} отсутствует обязательный лист 'Все КЦР'")
            ws = wb["Все КЦР"]
            ws_f = wb_f["Все КЦР"]
            h_row = detect_header_row(ws)

            expected_headers: Dict[int, List[str]] = {
                6: ["статус", "21.07.26"],
                7: ["стаус", "21.05.26"],
                10: ["действует"],
                11: ["блокировка"],
            }
            if ws.cell(h_row, 6).value is None:
                expected_headers.pop(6, None)
            if ws.cell(h_row, 7).value is None:
                expected_headers.pop(7, None)

            for c_tot, c_uns, c_pct, _, _, _, p_tok, d_tok in cls.SLICE_SPECS:
                expected_headers[c_tot] = ["всего", p_tok]
                expected_headers[c_uns] = ["долг", p_tok, d_tok]
                expected_headers[c_pct] = ["%", p_tok, d_tok]

            verify_header_tokens(ws, h_row, expected_headers, filename, "Все КЦР")

            clients: Dict[str, ClientRecord] = {}
            slices: List[SliceRecord] = []
            history: List[HistoryRecord] = []
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
                    c_rec = ClientRecord(
                        source_client_id=cid,
                        is_active=None,
                        is_blocked=None,
                        time_zone_offset=None,
                        admin_name=None,
                        provenance=Provenance(source_file=filename, source_sheet="Все КЦР", source_row=r)
                    )
                    set_client_attr(c_rec, "kcr_code", clean_optional_str(ws.cell(r, 2).value), filename, "Все КЦР", r, "B")
                    set_client_attr(c_rec, "organization_name", clean_optional_str(ws.cell(r, 3).value), filename, "Все КЦР", r, "C")
                    set_client_attr(c_rec, "service_center", clean_optional_str(ws.cell(r, 4).value), filename, "Все КЦР", r, "D")
                    set_client_attr(c_rec, "inn", clean_optional_str(ws.cell(r, 8).value), filename, "Все КЦР", r, "H")
                    set_client_attr(c_rec, "kpp", clean_optional_str(ws.cell(r, 9).value), filename, "Все КЦР", r, "I")
                    set_client_attr(c_rec, "is_active", clean_optional_str(ws.cell(r, 10).value), filename, "Все КЦР", r, "J")
                    set_client_attr(c_rec, "is_blocked", clean_optional_str(ws.cell(r, 11).value), filename, "Все КЦР", r, "K")
                    mgr_l = clean_optional_str(ws.cell(r, 12).value)
                    mgr_m = clean_optional_str(ws.cell(r, 13).value)
                    if mgr_l:
                        set_client_attr(c_rec, "manager_name", mgr_l, filename, "Все КЦР", r, "L")
                    elif mgr_m:
                        set_client_attr(c_rec, "manager_name", mgr_m, filename, "Все КЦР", r, "M")
                    clients[cid] = c_rec

                stat_f = clean_optional_str(ws.cell(r, 6).value)
                if stat_f:
                    history.append(HistoryRecord(
                        source_client_id=cid,
                        event_type="historical_status",
                        occurred_at="2026-07-21",
                        text=None,
                        raw_status=stat_f,
                        next_contact_date=None,
                        raw_deadline=None,
                        provenance=Provenance(source_file=filename, source_sheet="Все КЦР", source_row=r, source_column="F")
                    ))

                stat_g = clean_optional_str(ws.cell(r, 7).value)
                if stat_g:
                    history.append(HistoryRecord(
                        source_client_id=cid,
                        event_type="historical_status",
                        occurred_at="2026-05-21",
                        text=None,
                        raw_status=stat_g,
                        next_contact_date=None,
                        raw_deadline=None,
                        provenance=Provenance(source_file=filename, source_sheet="Все КЦР", source_row=r, source_column="G")
                    ))

                for c_tot, c_uns, c_pct, p_start, p_end, as_of, _, _ in cls.SLICE_SPECS:
                    rec = parse_slice_cell_triplet(
                        ws, r, c_tot, c_uns, c_pct, cid, p_start, p_end, as_of, filename, "Все КЦР", issues, ws_f=ws_f
                    )
                    if rec is not None:
                        slices.append(rec)

            return clients, slices, history, issues
        finally:
            wb.close()
            wb_f.close()


class PeriodicSliceParser:
    SHEET_PERIODS = {
        "весь 2025 год": ("2025-01-01", "2025-12-31"),
        "1 кв 2026": ("2026-01-01", "2026-03-31"),
        "2 кв 2026": ("2026-04-01", "2026-06-30"),
        "3 кв 2026": ("2026-07-01", "2026-09-30")
    }

    @classmethod
    def parse(
        cls,
        filepath: str,
        known_clients: Set[str],
        require_all_sheets: bool = False,
        clients_dict: Optional[Dict[str, ClientRecord]] = None
    ) -> Tuple[List[SliceRecord], List[QualityIssue]]:
        filename = os.path.basename(filepath)
        wb = openpyxl.load_workbook(filepath, data_only=True)
        wb_f = openpyxl.load_workbook(filepath, data_only=False)
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
                ws_f = wb_f[sname]
                h_row = detect_header_row(ws)
                col1_h = str(ws.cell(h_row, 1).value or "").lower()
                id_token = ["id"] if "id" in col1_h else ["идентификатор"]
                verify_header_tokens(ws, h_row, {
                    1: id_token,
                    4: ["всего"],
                    5: ["не отправлено"]
                }, filename, sname)

                col2_h = str(ws.cell(h_row, 2).value or "").lower()
                col3_h = str(ws.cell(h_row, 3).value or "").lower()

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
                    elif clients_dict is not None and cid in clients_dict:
                        c_rec = clients_dict[cid]
                        val_b = clean_optional_str(ws.cell(r, 2).value)
                        val_c = clean_optional_str(ws.cell(r, 3).value)
                        if "сц" in col2_h and "кцр" not in col2_h:
                            set_client_attr(c_rec, "service_center", val_b, filename, sname, r, "B")
                        else:
                            set_client_attr(c_rec, "kcr_code", val_b, filename, sname, r, "B")
                        if "сц" in col3_h:
                            set_client_attr(c_rec, "service_center", val_c, filename, sname, r, "C")
                        else:
                            set_client_attr(c_rec, "organization_name", val_c, filename, sname, r, "C")

                    rec = parse_slice_cell_triplet(
                        ws, r, 4, 5, None, cid, p_start, p_end, "2026-10-20",
                        filename, sname, issues, base_rejection=base_rejection, ws_f=ws_f
                    )
                    if rec is not None:
                        slices.append(rec)

            return slices, issues
        finally:
            wb.close()
            wb_f.close()


class HistoricalFilesParser:
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
        wb_f = openpyxl.load_workbook(filepath, data_only=False)
        try:
            if "1-3 кв 25" not in wb.sheetnames:
                raise ValueError(f"В файле {filename} отсутствует обязательный лист '1-3 кв 25'")
            ws = wb["1-3 кв 25"]
            ws_f = wb_f["1-3 кв 25"]
            h_row = detect_header_row(ws)

            expected_headers: Dict[int, List[str]] = {
                1: ["идентификатор"],
                10: ["претензии"],
                19: ["документов всего", "весь 2025"],
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
                    c_rec = clients_dict[cid]
                    set_client_attr(c_rec, "kcr_code", clean_optional_str(ws.cell(r, 2).value), filename, "1-3 кв 25", r, "B")
                    set_client_attr(c_rec, "service_center", clean_optional_str(ws.cell(r, 3).value), filename, "1-3 кв 25", r, "C")
                    cdate_d = parse_iso_date(ws.cell(r, 4).value) or clean_optional_str(ws.cell(r, 4).value)
                    set_client_attr(c_rec, "creation_date", cdate_d, filename, "1-3 кв 25", r, "D")
                    set_client_attr(c_rec, "time_zone_offset", clean_optional_str(ws.cell(r, 5).value), filename, "1-3 кв 25", r, "E")
                    set_client_attr(c_rec, "admin_name", clean_optional_str(ws.cell(r, 6).value), filename, "1-3 кв 25", r, "F")
                    set_client_attr(c_rec, "admin_phone", clean_optional_str(ws.cell(r, 7).value), filename, "1-3 кв 25", r, "G")
                    set_client_attr(c_rec, "admin_email", clean_optional_str(ws.cell(r, 8).value), filename, "1-3 кв 25", r, "H")

                for c_tot, c_uns, c_pct, as_of, is_ambig in hist_slice_specs:
                    base_rej = "UNKNOWN_CLIENT_ID" if is_unknown else (
                        "AMBIGUOUS_AS_OF_DATE:2025-01-16_vs_2026-01-16" if is_ambig else None
                    )
                    rec = parse_slice_cell_triplet(
                        ws, r, c_tot, c_uns, c_pct, cid, "2025-01-01", "2025-12-31", as_of,
                        filename, "1-3 кв 25", issues, base_rejection=base_rej, ws_f=ws_f
                    )
                    if rec is not None:
                        slices.append(rec)

                if is_unknown:
                    continue

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

                claim_comm = clean_optional_str(ws.cell(r, 10).value)
                if claim_comm:
                    history.append(HistoryRecord(
                        source_client_id=cid, event_type="claim_comment", occurred_at=None,
                        text=claim_comm, raw_status=None, next_contact_date=None, raw_deadline=None,
                        provenance=Provenance(source_file=filename, source_sheet="1-3 кв 25", source_row=r, source_column="J")
                    ))

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
            wb_f.close()

    @classmethod
    def parse_q1_2026(
        cls,
        filepath: str,
        known_clients: Optional[Set[str]] = None,
        clients_dict: Optional[Dict[str, ClientRecord]] = None
    ) -> Tuple[List[SliceRecord], List[HistoryRecord], List[QualityIssue]]:
        filename = os.path.basename(filepath)
        wb = openpyxl.load_workbook(filepath, data_only=True)
        wb_f = openpyxl.load_workbook(filepath, data_only=False)
        try:
            if "Все КЦР" not in wb.sheetnames:
                raise ValueError(f"В файле {filename} отсутствует обязательный лист 'Все КЦР'")
            ws = wb["Все КЦР"]
            ws_f = wb_f["Все КЦР"]
            h_row = detect_header_row(ws)

            expected_headers: Dict[int, List[str]] = {
                1: ["идентификатор"],
                15: ["стаус", "21.05.26"],
                22: ["всего", "весь 2025"],
                23: ["долг", "весь 2025", "21.05.26"],
                24: ["%", "весь 2025", "21.05.26"],
                25: ["всего", "2-4 кв 2025"],
                26: ["долг", "2-4 кв 2025", "21.05.26"],
                27: ["%", "2-4 кв 2025", "21.05.26"],
                28: ["всего", "1 кв 2026"],
                29: ["долг", "1 кв 2026", "21.05.26"],
                30: ["%", "1 кв 2026", "21.05.26"],
                31: ["всего", "2 кв 2025-1 кв 2026"],
                32: ["долг", "2 кв 2025-1 кв 2026", "21.05.26"],
                33: ["%", "2 кв 2025-1 кв 2026", "21.05.26"],
                35: ["комментарии оовс"],
                36: ["статус звонка"],
            }
            if ws.cell(h_row, 15).value is None:
                expected_headers.pop(15, None)

            verify_header_tokens(ws, h_row, expected_headers, filename, "Все КЦР")

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
                    c_rec = clients_dict[cid]
                    set_client_attr(c_rec, "kcr_code", clean_optional_str(ws.cell(r, 2).value), filename, "Все КЦР", r, "B")
                    set_client_attr(c_rec, "organization_name", clean_optional_str(ws.cell(r, 4).value), filename, "Все КЦР", r, "D")
                    set_client_attr(c_rec, "service_center", clean_optional_str(ws.cell(r, 5).value), filename, "Все КЦР", r, "E")
                    set_client_attr(c_rec, "time_zone_offset", clean_optional_str(ws.cell(r, 10).value), filename, "Все КЦР", r, "J")
                    set_client_attr(c_rec, "admin_name", clean_optional_str(ws.cell(r, 11).value), filename, "Все КЦР", r, "K")
                    set_client_attr(c_rec, "admin_phone", clean_optional_str(ws.cell(r, 12).value), filename, "Все КЦР", r, "L")
                    set_client_attr(c_rec, "admin_email", clean_optional_str(ws.cell(r, 13).value), filename, "Все КЦР", r, "M")
                    set_client_attr(c_rec, "inn", clean_optional_str(ws.cell(r, 16).value), filename, "Все КЦР", r, "P")
                    set_client_attr(c_rec, "kpp", clean_optional_str(ws.cell(r, 17).value), filename, "Все КЦР", r, "Q")
                    set_client_attr(c_rec, "is_active", clean_optional_str(ws.cell(r, 18).value), filename, "Все КЦР", r, "R")
                    set_client_attr(c_rec, "is_blocked", clean_optional_str(ws.cell(r, 19).value), filename, "Все КЦР", r, "S")
                    mgr_t = clean_optional_str(ws.cell(r, 20).value)
                    mgr_u = clean_optional_str(ws.cell(r, 21).value)
                    if mgr_t:
                        set_client_attr(c_rec, "manager_name", mgr_t, filename, "Все КЦР", r, "T")
                    elif mgr_u:
                        set_client_attr(c_rec, "manager_name", mgr_u, filename, "Все КЦР", r, "U")

                for c_tot, c_uns, c_pct, p_start, p_end in q1_slice_specs:
                    base_rej = "UNKNOWN_CLIENT_ID" if is_unknown else None
                    rec = parse_slice_cell_triplet(
                        ws, r, c_tot, c_uns, c_pct, cid, p_start, p_end, "2026-05-21",
                        filename, "Все КЦР", issues, base_rejection=base_rej, ws_f=ws_f
                    )
                    if rec is not None:
                        slices.append(rec)

                if is_unknown:
                    continue

                stat_o = clean_optional_str(ws.cell(r, 15).value)
                if stat_o:
                    history.append(HistoryRecord(
                        source_client_id=cid,
                        event_type="historical_status",
                        occurred_at="2026-05-21",
                        text=None,
                        raw_status=stat_o,
                        next_contact_date=None,
                        raw_deadline=None,
                        provenance=Provenance(source_file=filename, source_sheet="Все КЦР", source_row=r, source_column="O")
                    ))

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
            wb_f.close()


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

        # Предварительные опоры: только корректные записи с total, без конфликтов
        # одного ключа. После общего аудита опоры проверяются повторно.
        known_period_totals: Dict[Tuple[str, str, str], Set[Tuple[str, int]]] = {}
        reference_groups: Dict[Tuple[str, str, str, str], List[SliceRecord]] = {}
        for s in slices:
            if (
                s.rejection_reason is None
                and s.total is not None
                and s.total >= 0
                and (s.unsent is None or (0 <= s.unsent <= s.total))
                and is_valid_iso_date(s.period_start)
                and is_valid_iso_date(s.period_end)
                and is_valid_iso_date(s.as_of_date)
                and s.period_start <= s.period_end
                and (s.raw_percentage is None or math.isfinite(s.raw_percentage))
            ):
                key = (s.source_client_id, s.period_start, s.period_end, s.as_of_date)
                reference_groups.setdefault(key, []).append(s)
        for key, group in reference_groups.items():
            if len({(s.total, s.unsent) for s in group}) == 1:
                known_period_totals.setdefault(key[:3], set()).add((key[3], group[0].total))

        def validate_missing_total(s: SliceRecord, references) -> None:
            """Проверяет повторное наблюдение, сохраняя исходный total=None."""
            p = s.provenance
            pk = (s.source_client_id, s.period_start, s.period_end)
            other_date_totals = {tot for dt, tot in references.get(pk, set()) if dt != s.as_of_date}
            if len(other_date_totals) != 1:
                code = "UNSENT_WITHOUT_TOTAL"
                raw_value = f"total=null, unsent={s.unsent}"
                description = "Нет однозначного общего количества из принятых наблюдений периода на другую дату"
            elif s.unsent > next(iter(other_date_totals)):
                code = "UNSENT_EXCEEDS_TOTAL"
                raw_value = f"unsent={s.unsent}, period_total={next(iter(other_date_totals))}"
                description = "Остаток повторного среза без total превышает известное общее количество документов периода"
            else:
                return
            issues.append(QualityIssue(
                code=code, file=p.source_file, sheet=p.source_sheet, row=p.source_row,
                col=p.source_column or "total", raw_value=raw_value,
                description=description, action="requires_review"
            ))
            s.rejection_reason = code

        for s in slices:
            p = s.provenance

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
                validate_missing_total(s, known_period_totals)

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

            if s.rejection_reason:
                rejected.append(s)
                continue

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

        # Конфликт с принятым повтором без total тоже может удалить опору.
        # В итоговом valid не должно остаться записей, зависящих от неё.
        accepted_period_totals: Dict[Tuple[str, str, str], Set[Tuple[str, int]]] = {}
        for key, s in valid_by_key.items():
            if s.total is not None:
                accepted_period_totals.setdefault(key[:3], set()).add((s.as_of_date, s.total))
        for key, s in list(valid_by_key.items()):
            if s.total is None and s.unsent is not None:
                validate_missing_total(s, accepted_period_totals)
                if s.rejection_reason:
                    valid_by_key.pop(key)
                    rejected.append(s)

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

    clients_dict, s1, h1, issues1 = BaseRegistryParser.parse(f1)
    known_cids = set(clients_dict.keys())

    s2, issues2 = PeriodicSliceParser.parse(f2, known_cids, require_all_sheets=True, clients_dict=clients_dict)
    s3, h3, issues3 = HistoricalFilesParser.parse_debts_2025(f3, known_clients=known_cids, clients_dict=clients_dict)
    s4, h4, issues4 = HistoricalFilesParser.parse_q1_2026(f4, known_clients=known_cids, clients_dict=clients_dict)

    all_raw_slices = s1 + s2 + s3 + s4
    valid_slices, rejected_slices, audit_issues = QualityAuditor.audit_and_filter_slices(all_raw_slices)

    all_history = h1 + h3 + h4
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

    # Полное восстановление предыдущего комплекта целиком при прерванном запуске
    if os.path.exists(backup_dir):
        for fname in target_files:
            bck_file = os.path.join(backup_dir, fname)
            dst_file = os.path.join(output_dir, fname)
            if os.path.exists(bck_file):
                os.replace(bck_file, dst_file)
        shutil.rmtree(backup_dir, ignore_errors=True)

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


def find_source_dir() -> str:
    req = "КЦР_последняя выгрузка_Дубна.xlsx"
    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [os.path.join(script_dir, "sources"), ".", "..", "../..", script_dir, os.path.join(script_dir, ".."), os.path.join(script_dir, "../..")]
    for c in candidates:
        if os.path.exists(os.path.join(c, req)):
            return c
    return "."


def main():
    base_dir = find_source_dir()
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
