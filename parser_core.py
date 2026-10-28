"""
parser_core.py — Модуль разбора Excel и нормализации в единый JSON-контракт
Контур КЦР. Задача 1: Разбор Excel и единый формат.
"""

import os
import json
import csv
from datetime import datetime, date
import openpyxl

def format_val(val):
    """Форматирование значений ячеек: даты в YYYY-MM-DD, пустые строки в None."""
    if val is None:
        return None
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, str):
        v = val.strip()
        return v if v != "" else None
    return val

def parse_all_kcr(base_dir="."):
    """
    Разбор 4 Excel-файлов, извлечение клиентов, срезов, истории и аудит данных.
    Возвращает: clients (dict/list), slices (list), history (list), errors (list).
    """
    clients = {}
    slices = []
    history = []
    errors = []

    # 1. КЦР_последняя выгрузка_Дубна.xlsx
    file1 = os.path.join(base_dir, "КЦР_последняя выгрузка_Дубна.xlsx")
    if os.path.exists(file1):
        wb1 = openpyxl.load_workbook(file1, data_only=True)
        ws1 = wb1["Все КЦР"]

        slice_specs = [
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

        for r in range(3, ws1.max_row + 1):
            cid = ws1.cell(r, 1).value
            if cid is None:
                if any(ws1.cell(r, c).value is not None for c in range(1, ws1.max_column + 1)):
                    errors.append({
                        "code": "MISSING_CLIENT_ID",
                        "file": "КЦР_последняя выгрузка_Дубна.xlsx",
                        "sheet": "Все КЦР",
                        "row": r, "col": "A",
                        "raw_value": None,
                        "description": "Служебная или итоговая строка без идентификатора клиента",
                        "action": "skipped"
                    })
                continue

            cid_str = str(cid).strip()
            if cid_str in clients:
                errors.append({
                    "code": "DUPLICATE_CLIENT_ID",
                    "file": "КЦР_последняя выгрузка_Дубна.xlsx",
                    "sheet": "Все КЦР",
                    "row": r, "col": "A",
                    "raw_value": cid_str,
                    "description": f"Повторяющийся идентификатор клиента: {cid_str}",
                    "action": "requires_review"
                })
            else:
                clients[cid_str] = {
                    "source_client_id": cid_str,
                    "is_active": format_val(ws1.cell(r, 10).value),
                    "is_blocked": format_val(ws1.cell(r, 11).value),
                    "provenance": {
                        "source_file": "КЦР_последняя выгрузка_Дубна.xlsx",
                        "source_sheet": "Все КЦР",
                        "source_row": r
                    }
                }

            for c_tot, c_uns, c_pct, p_start, p_end, as_of in slice_specs:
                tot = ws1.cell(r, c_tot).value
                uns = ws1.cell(r, c_uns).value
                pct = ws1.cell(r, c_pct).value
                if tot is not None or uns is not None:
                    slices.append({
                        "source_client_id": cid_str,
                        "period_start": p_start,
                        "period_end": p_end,
                        "as_of_date": as_of,
                        "total": int(tot) if isinstance(tot, (int, float)) and not isinstance(tot, bool) else tot,
                        "unsent": int(uns) if isinstance(uns, (int, float)) and not isinstance(uns, bool) else uns,
                        "raw_percentage": pct,
                        "provenance": {
                            "source_file": "КЦР_последняя выгрузка_Дубна.xlsx",
                            "source_sheet": "Все КЦР",
                            "source_row": r,
                            "source_column": f"{openpyxl.utils.get_column_letter(c_tot)},{openpyxl.utils.get_column_letter(c_uns)}"
                        }
                    })

    # 2. Срез 20.10.26_Дубна.xlsx
    file2 = os.path.join(base_dir, "Срез 20.10.26_Дубна.xlsx")
    if os.path.exists(file2):
        wb2 = openpyxl.load_workbook(file2, data_only=True)
        sheet_periods = {
            "весь 2025 год": ("2025-01-01", "2025-12-31"),
            "1 кв 2026": ("2026-01-01", "2026-03-31"),
            "2 кв 2026": ("2026-04-01", "2026-06-30"),
            "3 кв 2026": ("2026-07-01", "2026-09-30")
        }
        seen_slice_ids = {s: set() for s in sheet_periods}

        for sname, (p_start, p_end) in sheet_periods.items():
            ws2 = wb2[sname]
            for r in range(3, ws2.max_row + 1):
                cid = ws2.cell(r, 1).value
                if cid is None:
                    continue
                cid_str = str(cid).strip()

                if cid_str in seen_slice_ids[sname]:
                    errors.append({
                        "code": "DUPLICATE_CLIENT_ID",
                        "file": "Срез 20.10.26_Дубна.xlsx",
                        "sheet": sname,
                        "row": r, "col": "A",
                        "raw_value": cid_str,
                        "description": f"Дублирующийся идентификатор в рамках одного листа: {cid_str}",
                        "action": "requires_review"
                    })
                else:
                    seen_slice_ids[sname].add(cid_str)

                if cid_str not in clients:
                    errors.append({
                        "code": "UNKNOWN_CLIENT_ID",
                        "file": "Срез 20.10.26_Дубна.xlsx",
                        "sheet": sname,
                        "row": r, "col": "A",
                        "raw_value": cid_str,
                        "description": f"Клиент {cid_str} отсутствует в основной базе клиентов",
                        "action": "blocked"
                    })

                tot = ws2.cell(r, 4).value
                uns = ws2.cell(r, 5).value
                slices.append({
                    "source_client_id": cid_str,
                    "period_start": p_start,
                    "period_end": p_end,
                    "as_of_date": "2026-10-20",
                    "total": int(tot) if isinstance(tot, (int, float)) and not isinstance(tot, bool) else tot,
                    "unsent": int(uns) if isinstance(uns, (int, float)) and not isinstance(uns, bool) else uns,
                    "raw_percentage": None,
                    "provenance": {
                        "source_file": "Срез 20.10.26_Дубна.xlsx",
                        "source_sheet": sname,
                        "source_row": r
                    }
                })

    # 3. База должники III кв 2025 Дубна.xlsx
    file3 = os.path.join(base_dir, "База должники III кв 2025 Дубна.xlsx")
    if os.path.exists(file3):
        wb3 = openpyxl.load_workbook(file3, data_only=True)
        ws3 = wb3["1-3 кв 25"]
        for r in range(2, ws3.max_row + 1):
            cid = ws3.cell(r, 1).value
            if cid is None:
                if any(ws3.cell(r, c).value is not None for c in range(1, ws3.max_column + 1)):
                    errors.append({
                        "code": "MISSING_CLIENT_ID",
                        "file": "База должники III кв 2025 Дубна.xlsx",
                        "sheet": "1-3 кв 25",
                        "row": r, "col": "A",
                        "raw_value": None,
                        "description": "Служебная или итоговая строка без идентификатора клиента",
                        "action": "skipped"
                    })
                continue
            cid_str = str(cid).strip()

            comment_1 = format_val(ws3.cell(r, 46).value)
            comment_2 = format_val(ws3.cell(r, 53).value)
            claim_comm = format_val(ws3.cell(r, 10).value)
            call_status = format_val(ws3.cell(r, 47).value)
            call_date = format_val(ws3.cell(r, 48).value)
            client_deadline = format_val(ws3.cell(r, 49).value)

            full_text_parts = [p for p in [claim_comm, comment_1, comment_2] if p]
            full_text = " | ".join(full_text_parts) if full_text_parts else None

            if full_text or call_status or call_date or client_deadline:
                history.append({
                    "source_client_id": cid_str,
                    "event_type": "communication_log",
                    "occurred_at": call_date,
                    "text": full_text,
                    "raw_status": call_status,
                    "raw_deadline": client_deadline,
                    "provenance": {
                        "source_file": "База должники III кв 2025 Дубна.xlsx",
                        "source_sheet": "1-3 кв 25",
                        "source_row": r
                    }
                })

    # 4. База I кв 2026 Дубна.xlsx
    file4 = os.path.join(base_dir, "База I кв 2026 Дубна.xlsx")
    if os.path.exists(file4):
        wb4 = openpyxl.load_workbook(file4, data_only=True)
        ws4 = wb4["Все КЦР"]
        for r in range(3, ws4.max_row + 1):
            cid = ws4.cell(r, 1).value
            if cid is None:
                continue
            cid_str = str(cid).strip()

            comm_oovs = format_val(ws4.cell(r, 35).value)
            call_stat = format_val(ws4.cell(r, 36).value)
            recall_date = format_val(ws4.cell(r, 37).value)
            client_dl = format_val(ws4.cell(r, 38).value)

            if comm_oovs or call_stat or recall_date or client_dl:
                history.append({
                    "source_client_id": cid_str,
                    "event_type": "operator_contact",
                    "occurred_at": recall_date,
                    "text": comm_oovs,
                    "raw_status": call_stat,
                    "raw_deadline": client_dl,
                    "provenance": {
                        "source_file": "База I кв 2026 Дубна.xlsx",
                        "source_sheet": "Все КЦР",
                        "source_row": r
                    }
                })

    # Аудит срезов на расхождения и некорректные значения
    for item in slices:
        tot = item["total"]
        uns = item["unsent"]
        pct = item["raw_percentage"]
        prov = item["provenance"]

        if tot is not None and (not isinstance(tot, (int, float)) or tot < 0):
            errors.append({
                "code": "INVALID_NUMERIC_TOTAL",
                "file": prov["source_file"], "sheet": prov["source_sheet"], "row": prov["source_row"],
                "col": "total", "raw_value": str(tot),
                "description": "Общее количество документов не является неотрицательным числом",
                "action": "requires_review"
            })

        if uns is not None and (not isinstance(uns, (int, float)) or uns < 0):
            errors.append({
                "code": "INVALID_NUMERIC_UNSENT",
                "file": prov["source_file"], "sheet": prov["source_sheet"], "row": prov["source_row"],
                "col": "unsent", "raw_value": str(uns),
                "description": "Количество непереданных документов не является неотрицательным числом",
                "action": "requires_review"
            })

        if isinstance(tot, (int, float)) and isinstance(uns, (int, float)):
            if uns > tot:
                errors.append({
                    "code": "UNSENT_EXCEEDS_TOTAL",
                    "file": prov["source_file"], "sheet": prov["source_sheet"], "row": prov["source_row"],
                    "col": "unsent", "raw_value": f"unsent={uns}, total={tot}",
                    "description": "Количество непереданных документов превышает общее число документов",
                    "action": "requires_review"
                })

            if pct is not None and isinstance(pct, (int, float)) and tot > 0:
                calc_pct = uns / tot
                if abs(pct - calc_pct) > 0.02:
                    errors.append({
                        "code": "PERCENTAGE_MISMATCH",
                        "file": prov["source_file"], "sheet": prov["source_sheet"], "row": prov["source_row"],
                        "col": "percentage", "raw_value": f"excel={pct:.4f}, calc={calc_pct:.4f}",
                        "description": "Исходный процент расходится с вычисленным unsent/total",
                        "action": "requires_review"
                    })

    return list(clients.values()), slices, history, errors

def main():
    clients, slices, history, errors = parse_all_kcr()
    
    with open("normalized_clients.json", "w", encoding="utf-8") as f:
        json.dump(clients, f, ensure_ascii=False, indent=2)

    with open("normalized_slices.jsonl", "w", encoding="utf-8") as f:
        for s in slices:
            f.write(json.dumps(s, ensure_ascii=False) + "\n")

    with open("normalized_history.jsonl", "w", encoding="utf-8") as f:
        for h in history:
            f.write(json.dumps(h, ensure_ascii=False) + "\n")

    with open("quality_report.json", "w", encoding="utf-8") as f:
        json.dump(errors, f, ensure_ascii=False, indent=2)

    with open("quality_report.csv", "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["code", "file", "sheet", "row", "col", "raw_value", "description", "action"])
        writer.writeheader()
        for err in errors:
            writer.writerow(err)

    print(f"Парсинг успешно завершён!")
    print(f"Клиентов: {len(clients)}")
    print(f"Срезов: {len(slices)}")
    print(f"Исторических записей: {len(history)}")
    print(f"Замечаний качества: {len(errors)}")

if __name__ == "__main__":
    main()
