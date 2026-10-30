# Контракт обмена v8

Шесть основных наборов сохраняют структуру v7. Дополнительные файлы проверки описаны в README.md.

## Контракт обмена для участников Задач №2–4

Функция `parse_all_kcr(base_dir)` возвращает кортеж из **5 коллекций**:
`(clients, valid_slices, rejected_slices, history, quality_issues)`.

Поля `ClientRecord` (`normalized_clients.json`):
- `source_client_id`: `str`
- `is_active`: `str | null`
- `is_blocked`: `str | null`
- `time_zone_offset`: `str | null`
- `admin_name`: `str | null` (ФИО главного администратора из выделенных колонок `F` / `K`)
- `kcr_code`: `str | null`
- `organization_name`: `str | null`
- `service_center`: `str | null`
- `inn`: `str | null`
- `kpp`: `str | null`
- `creation_date`: `str | null` (`YYYY-MM-DD`)
- `admin_phone`: `str | null`
- `admin_email`: `str | null`
- `manager_name`: `str | null` (ФИО менеджера из колонок `L`/`M` и `T`/`U`)
- `provenance`: `{"source_file", "source_sheet", "source_row", "source_column"}`
- `attribute_provenance`: словарь источников для каждого заполненного или обогащённого атрибута карточки

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

