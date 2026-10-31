# Контракт обмена v9

По сравнению с v8 в ClientRecord добавлены два необязательных строковых поля: `service_center_code` и `raw_creation_date`. Остальные поля сохранены; `service_center` теперь однозначно означает название СЦ. В текущих исходниках оба новых поля пусты. Дополнительные файлы проверки описаны в README.md.

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
- `service_center`: `str | null` (название СЦ)
- `service_center_code`: `str | null` (код СЦ, отдельно от названия и kcr_code)
- `inn`: `str | null`
- `kpp`: `str | null`
- `creation_date`: `str | null` (только календарная дата `YYYY-MM-DD`)
- `raw_creation_date`: `str | null` (исходное значение; строка сохраняется без удаления внешних пробелов)
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


## Изменения в проверках v9

Нераспознанная дата создания не попадает в creation_date. Исходное значение сохраняется в raw_creation_date с собственным attribute_provenance, в отчёт добавляется INVALID_CREATION_DATE. При пустом исходном значении оба поля остаются null, ошибка не создаётся. При распознанном значении сохраняются календарная дата и исходное представление.

На квартальных листах «код СЦ» из B переносится в service_center_code, «название СЦ» из C — в service_center. На годовом листе «Код КЦР» остаётся kcr_code. Заполненные реквизиты по-прежнему не перезаписываются следующими источниками, происхождение сохраняется отдельно для каждого поля.

Для принятого повтора с total=null процент сверяется по однозначному total окончательно принятых опор того же клиента и периода на другую дату. Порог расхождения прежний: больше 0.02. PERCENTAGE_MISMATCH требует проверки, но не отклоняет срез. Исходные total=null и raw_percentage не заменяются вычисленными значениями. При total=0 процент по делению не проверяется.
