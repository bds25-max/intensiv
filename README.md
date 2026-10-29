# Контур КЦР — Модуль разбора Excel и нормализации данных (Задача 1)

Автономный ETL-модуль для разбора четырёх таблиц Excel, извлечения базы клиентов, нормализации срезов документов, сохранения полной структурированной истории статусов и контактов (без какого-либо маскирования и искусственного обезличивания) и формирования отчёта качества данных.

---

## 📌 Ключевые характеристики актуальной сборки

- **5 874 клиента** (`normalized_clients.json`) — карточки дополнены доступными в исходниках полями (`time_zone_offset` из столбца `J` базы I кв. 2026).
- **87 933 допустимых среза** (`normalized_slices.jsonl`) — полностью очищены от неизвестных клиентов, дублей, конфликтов, неоднозначных дат, `NaN`/`Infinity` и числовых ошибок.
- **3 877 отклонённых срезов** (`rejected_slices.jsonl`) — изолированные записи с явно указанной причиной `rejection_reason`.
- **3 029 записей истории** (`normalized_history.jsonl`) — все тексты, комментарии, номера телефонов, email и ФИО сохранены дословно:
  - **2 280 исторических статусов** (`event_type="historical_status"`) из колонок `K–Q` и `AO–AP` с датой `occurred_at`;
  - **52 комментария для претензии** (столбец `J`, `event_type="claim_comment"`);
  - **295 записей первого касания** (`AT–AX`, `event_type="first_contact"`) с раздельным сохранением `raw_send_deadline` и `raw_client_date`;
  - **280 записей повторного звонка** (`BA–BE`, `event_type="second_contact"`);
  - **7 записей инцидентов и запросов эксперта** (`BF–BH`, `event_type="incident_or_expert_request"`);
  - **115 записей контактов I кв. 2026** (`AI–AM`, `event_type="operator_contact"`).
- **3 683 замечания качества** (`quality_report.json` и `quality_report.csv`).
- **Защищённый транзакционный экспорт**: резервная копия `_backup` сохраняется на диске до подтверждённого завершения публикации или полного восстановления при сбое.

---

## 📋 Контракт обмена для участников Задач №2–4

Функция `parse_all_kcr(base_dir)` возвращает кортеж из **5 коллекций**:
`(clients, valid_slices, rejected_slices, history, quality_issues)`.

Поля `HistoryRecord` (`normalized_history.jsonl`):
- `source_client_id`: `str`
- `event_type`: `'historical_status' | 'claim_comment' | 'first_contact' | 'second_contact' | 'incident_or_expert_request' | 'operator_contact'`
- `occurred_at`: `str | null` (`YYYY-MM-DD`, заполнено для `historical_status`)
- `text`: `str | null` (исходный текст без цензуры и маскирования)
- `raw_status`: `str | null`
- `next_contact_date`: `str | null` (`YYYY-MM-DD`)
- `raw_deadline`: `str | null`
- `raw_next_contact`: `str | null`
- `raw_send_deadline`: `str | null`
- `raw_client_date`: `str | null`
- `provenance`: `{"source_file", "source_sheet", "source_row", "source_column"}`

---

## 🚀 Быстрый старт и проверка

```bash
pip install -r requirements.txt
python parser_core.py
python -m unittest test_parser.py -v
```
