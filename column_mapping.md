# Карта колонок четырёх источников

Проверено по приложенным Excel: 4 файла, 7 листов, 188 колонок. Число заполненных ячеек считается ниже строки заголовков, включая итоговые/служебные строки без ID. Координаты таких строк сохраняются в отчёте качества или дополнительном исходном наборе.

## Правила преобразования

Идентификатор берётся из A. Пропуски сохраняются как null. Количества разбираются строго, исходные проценты сохраняются для проверки. Период и дата наблюдения задаются раздельно по явной схеме колонок. Менеджер и главный администратор — отдельные поля.

Повтор без total допускается только при однозначном количестве в принятых наблюдениях того же клиента и периода на другую дату. Количество не подставляется в пустую исходную ячейку. Это техническое правило текущей реализации, а не подтверждение бизнес-правила заказчиком.

Колонки, не перенесённые в основные схемы, перечислены явно. Их непустые исходные значения, а также исходные ячейки истории сохранены в source_preserved_values.jsonl. Это набор исходных ячеек с координатами; он не назначает неизвестным полям смысл события и не заменяет согласование схемы. Сохранены формулы и их доступные кешированные значения. Основные текстовые поля используют прежнее удаление внешних пробелов; дополнительный набор сохраняет исходные строки с пробелами.

Карта данных загружается из column_mapping.json. В v9 код СЦ выделен в service_center_code; service_center означает название. Дата создания проходит календарную проверку; исходное представление хранится в raw_creation_date. Процент повтора с total=null проверяется по окончательно принятой опоре, если её количество больше нуля.

## КЦР_последняя выгрузка_Дубна.xlsx — Все КЦР

Строка заголовков: 2. Колонок: 66.

| Колонка | Заголовок Excel | Непустых ячеек | Основное поле | Преобразование | Причина исключения |
| --- | --- | ---: | --- | --- | --- |
| A | (без заголовка) | 5874 | ClientRecord.source_client_id | str.strip(); пустые пропускаются, дубли фиксируются | — |
| B | Код КЦР | 0 | ClientRecord.kcr_code | clean_optional_str (в обезличенной выборке пусто) | — |
| C | Название | 0 | ClientRecord.organization_name | clean_optional_str (в обезличенной выборке пусто) | — |
| D | СЦ | 0 | ClientRecord.service_center | clean_optional_str (в обезличенной выборке пусто) | — |
| E | (без заголовка) | 0 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Нет непустых значений ниже заголовка в предоставленном файле. |
| F | Статус на 21.07.26 | 7 | HistoryRecord(historical_status) | occurred_at='2026-07-21', source_column='F' (7 записей) | — |
| G | стаус на 21.05.26 | 2748 | HistoryRecord(historical_status) | occurred_at='2026-05-21', source_column='G' (2 748 записей) | — |
| H | ИНН | 0 | ClientRecord.inn | clean_optional_str (в обезличенной выборке пусто) | — |
| I | КПП | 0 | ClientRecord.kpp | clean_optional_str (в обезличенной выборке пусто) | — |
| J | Действует ли КЦР | 5873 | ClientRecord.is_active | clean_optional_str (5 873 значения) | — |
| K | Блокировка в сервисе | 5873 | ClientRecord.is_blocked | clean_optional_str (5 873 значения) | — |
| L | ФИО менеджера | 0 | ClientRecord.manager_name | clean_optional_str; сохраняется отдельно от admin_name | — |
| M | ФИО менеджера из КП | 0 | ClientRecord.manager_name | Резервное поле manager_name, если L пусто | — |
| N | Статус на 21.07.26 | 0 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Нет непустых значений ниже заголовка в предоставленном файле. |
| O | Всего документов весь 2025 | 3419 | SliceRecord.total | parse_strict_int; период 2025-01-01..2025-12-31 | — |
| P | Долг весь 2025 на 29.09.26 | 3419 | SliceRecord.unsent | parse_strict_int; период 2025-01-01..2025-12-31, срез 2026-09-29 | — |
| Q | % Долга весь 2025 на 29.09.26 | 3419 | SliceRecord.raw_percentage | parse_strict_float; сверка с unsent/total (±0.02), срез 2026-09-29 | — |
| R | Долг весь 2025 на 21.07.26 | 3419 | SliceRecord.unsent | parse_strict_int; период 2025-01-01..2025-12-31, срез 2026-07-21 | — |
| S | % Долга весь 2025 на 21.07.26 | 3419 | SliceRecord.raw_percentage | parse_strict_float; сверка с unsent/total (±0.02), срез 2026-07-21 | — |
| T | Долг весь 2025 на 07.07.26 | 3419 | SliceRecord.unsent | parse_strict_int; период 2025-01-01..2025-12-31, срез 2026-07-07 | — |
| U | % Долга весь 2025 на 07.07.26 | 3418 | SliceRecord.raw_percentage | parse_strict_float; сверка с unsent/total (±0.02), срез 2026-07-07 | — |
| V | Долг весь 2025 на 21.05.26 | 3418 | SliceRecord.unsent | parse_strict_int; период 2025-01-01..2025-12-31, срез 2026-05-21 | — |
| W | % Долга весь 2025 на 21.05.26 | 3418 | SliceRecord.raw_percentage | parse_strict_float; сверка с unsent/total (±0.02), срез 2026-05-21 | — |
| X | Всего документов 2-4 кв 2025 | 3298 | SliceRecord.total | parse_strict_int; период 2025-04-01..2025-12-31 | — |
| Y | Долг 2-4 кв 2025 на 21.07.26 | 3298 | SliceRecord.unsent | parse_strict_int; период 2025-04-01..2025-12-31, срез 2026-07-21 | — |
| Z | % Долга 2-4кв 2025 на 21.07.26 | 3298 | SliceRecord.raw_percentage | parse_strict_float; сверка с unsent/total (±0.02), срез 2026-07-21 | — |
| AA | Долг 2-4 кв 2025 на 07.07.26 | 3298 | SliceRecord.unsent | parse_strict_int; период 2025-04-01..2025-12-31, срез 2026-07-07 | — |
| AB | % Долга 2-4кв 2025 на 07.07.26 | 3297 | SliceRecord.raw_percentage | parse_strict_float; сверка с unsent/total (±0.02), срез 2026-07-07 | — |
| AC | Долг 2-4 кв 2025 на 21.05.26 | 3297 | SliceRecord.unsent | parse_strict_int; период 2025-04-01..2025-12-31, срез 2026-05-21 | — |
| AD | % Долга 2-4кв 2025 на 21.05.26 | 3297 | SliceRecord.raw_percentage | parse_strict_float; сверка с unsent/total (±0.02), срез 2026-05-21 | — |
| AE | Всего документов 1 кв 2026 | 3181 | SliceRecord.total | parse_strict_int; период 2026-01-01..2026-03-31 | — |
| AF | Долг 1 кв 2026 на 29.09.26 | 3181 | SliceRecord.unsent | parse_strict_int; период 2026-01-01..2026-03-31, срез 2026-09-29 | — |
| AG | % Долга 1 кв 2026 на 29.09.26 | 3181 | SliceRecord.raw_percentage | parse_strict_float; сверка с unsent/total (±0.02), срез 2026-09-29 | — |
| AH | Долг 1 кв 2026 на 21.07.26 | 2928 | SliceRecord.unsent | parse_strict_int; период 2026-01-01..2026-03-31, срез 2026-07-21 | — |
| AI | % Долга 1 кв 2026 на 21.07.26 | 55 | SliceRecord.raw_percentage | parse_strict_float; сверка с unsent/total (±0.02), срез 2026-07-21 | — |
| AJ | Долг 1 кв 2026 на 07.07.26 | 2928 | SliceRecord.unsent | parse_strict_int; период 2026-01-01..2026-03-31, срез 2026-07-07 | — |
| AK | % Долга 1 кв 2026 на 07.07.26 | 2928 | SliceRecord.raw_percentage | parse_strict_float; сверка с unsent/total (±0.02), срез 2026-07-07 | — |
| AL | Долг 1 кв 2026 на 21.05.26 | 3180 | SliceRecord.unsent | parse_strict_int; период 2026-01-01..2026-03-31, срез 2026-05-21 | — |
| AM | % Долга 1 кв 2026 на 21.05.26 | 3180 | SliceRecord.raw_percentage | parse_strict_float; сверка с unsent/total (±0.02), срез 2026-05-21 | — |
| AN | Всего документов 2 кв 2025-1 кв 2026 | 3657 | SliceRecord.total | parse_strict_int; период 2025-04-01..2026-03-31 | — |
| AO | Долг 2 кв 2025-1 кв 2026 на 21.07.26 | 3657 | SliceRecord.unsent | parse_strict_int; период 2025-04-01..2026-03-31, срез 2026-07-21 | — |
| AP | % Долга 2 кв 2025-1 кв 2026 на 21.07.26 | 3657 | SliceRecord.raw_percentage | parse_strict_float; сверка с unsent/total (±0.02), срез 2026-07-21 | — |
| AQ | Долг 2 кв 2025-1 кв 2026 на 07.07.26 | 3657 | SliceRecord.unsent | parse_strict_int; период 2025-04-01..2026-03-31, срез 2026-07-07 | — |
| AR | % Долга 2 кв 2025-1 кв 2026 на 07.07.26 | 3657 | SliceRecord.raw_percentage | parse_strict_float; сверка с unsent/total (±0.02), срез 2026-07-07 | — |
| AS | Долг 2 кв 2025-1 кв 2026 на 21.05.26 | 5713 | SliceRecord.unsent | parse_strict_int; период 2025-04-01..2026-03-31, срез 2026-05-21 | — |
| AT | % Долга 2 кв 2025-1 кв 2026 на 21.05.26 | 3656 | SliceRecord.raw_percentage | parse_strict_float; сверка с unsent/total (±0.02), срез 2026-05-21 | — |
| AU | Всего документов 3-4 кв 2025 | 3153 | SliceRecord.total | parse_strict_int; период 2025-07-01..2025-12-31 | — |
| AV | Долг 3-4 кв 2025 на 29.09.26 | 3298 | SliceRecord.unsent | parse_strict_int; период 2025-07-01..2025-12-31, срез 2026-09-29 | — |
| AW | % Долг 3-4 кв 2025 на 29.09.26 | 3153 | SliceRecord.raw_percentage | parse_strict_float; сверка с unsent/total (±0.02), срез 2026-09-29 | — |
| AX | Долг 3-4 кв 2025 на 21.07.26 | 3153 | SliceRecord.unsent | parse_strict_int; период 2025-07-01..2025-12-31, срез 2026-07-21 | — |
| AY | % Долг 3-4 кв 2025 на 21.07.26 | 3153 | SliceRecord.raw_percentage | parse_strict_float; сверка с unsent/total (±0.02), срез 2026-07-21 | — |
| AZ | Всего документов 2 кв 2026 | 3021 | SliceRecord.total | parse_strict_int; период 2026-04-01..2026-06-30 | — |
| BA | Долг 2 кв 2026 на 29.09.26 | 3021 | SliceRecord.unsent | parse_strict_int; период 2026-04-01..2026-06-30, срез 2026-09-29 | — |
| BB | % Долг 2 кв 2026 на 29.09.26 | 3021 | SliceRecord.raw_percentage | parse_strict_float; сверка с unsent/total (±0.02), срез 2026-09-29 | — |
| BC | Долг 2 кв 2026 на 21.07.26 | 3021 | SliceRecord.unsent | parse_strict_int; период 2026-04-01..2026-06-30, срез 2026-07-21 | — |
| BD | % Долг 2 кв 2026 на 21.07.26 | 2591 | SliceRecord.raw_percentage | parse_strict_float; сверка с unsent/total (±0.02), срез 2026-07-21 | — |
| BE | Всего документов 3 кв 2025 - 2 кв 2026 | 3666 | SliceRecord.total | parse_strict_int; период 2025-07-01..2026-06-30 | — |
| BF | Долг по документам 3 кв 2025-2 кв 2026 на 29.09.26 | 3666 | SliceRecord.unsent | parse_strict_int; период 2025-07-01..2026-06-30, срез 2026-09-29 | — |
| BG | % Долг по документам 3 кв 2025-2 кв 2026 на 29.09.26 | 3666 | SliceRecord.raw_percentage | parse_strict_float; сверка с unsent/total (±0.02), срез 2026-09-29 | — |
| BH | Долг по документам 3 кв 2025-2 кв 2026 на 21.07.26 | 3666 | SliceRecord.unsent | parse_strict_int; период 2025-07-01..2026-06-30, срез 2026-07-21 | — |
| BI | % Долг по документам 3 кв 2025-2 кв 2026 на 21.07.26 | 3666 | SliceRecord.raw_percentage | parse_strict_float; сверка с unsent/total (±0.02), срез 2026-07-21 | — |
| BJ | Всего документов весь 2025 - 2 кв 2026 | 3846 | SliceRecord.total | parse_strict_int; период 2025-01-01..2026-06-30 | — |
| BK | Долг по документам  весь 2025 - 2 кв 2026 на 29.09.26 | 3846 | SliceRecord.unsent | parse_strict_int; период 2025-01-01..2026-06-30, срез 2026-09-29 | — |
| BL | % Долг по документам  весь 2025 - 2 кв 2026 на 29.09.26 | 3846 | SliceRecord.raw_percentage | parse_strict_float; сверка с unsent/total (±0.02), срез 2026-09-29 | — |
| BM | Долг по документам  весь 2025 - 2 кв 2026 на 21.07.26 | 3846 | SliceRecord.unsent | parse_strict_int; период 2025-01-01..2026-06-30, срез 2026-07-21 | — |
| BN | % Долг по документам  весь 2025 - 2 кв 2026 на 21.07.26 | 3846 | SliceRecord.raw_percentage | parse_strict_float; сверка с unsent/total (±0.02), срез 2026-07-21 | — |

## Срез 20.10.26_Дубна.xlsx — весь 2025 год

Строка заголовков: 2. Колонок: 5.

| Колонка | Заголовок Excel | Непустых ячеек | Основное поле | Преобразование | Причина исключения |
| --- | --- | ---: | --- | --- | --- |
| A | Уникальны обезличинный идентификатор | 3419 | SliceRecord.source_client_id | Сопоставление с белым списком; неизвестные -> UNKNOWN_CLIENT_ID (blocked) | — |
| B | Код КЦР | 0 | ClientRecord.kcr_code | clean_optional_str (обогащение карточки клиента при наличии) | — |
| C | Наитменование организации | 0 | ClientRecord.organization_name | clean_optional_str (обогащение карточки клиента при наличии) | — |
| D | документов всего | 3419 | SliceRecord.total | parse_strict_int; период 2025-01-01..2025-12-31, срез 2026-10-20 | — |
| E | Всего не отправлено документов | 3419 | SliceRecord.unsent | parse_strict_int; период 2025-01-01..2025-12-31, срез 2026-10-20 | — |

## Срез 20.10.26_Дубна.xlsx — 1 кв 2026

Строка заголовков: 2. Колонок: 5.

| Колонка | Заголовок Excel | Непустых ячеек | Основное поле | Преобразование | Причина исключения |
| --- | --- | ---: | --- | --- | --- |
| A | Уникальны обезличинный идентификатор | 2928 | SliceRecord.source_client_id | Сопоставление с белым списком; неизвестные -> UNKNOWN_CLIENT_ID (blocked) | — |
| B | код СЦ | 0 | ClientRecord.service_center_code | clean_optional_str; код СЦ отдельно от названия service_center; происхождение B | — |
| C | название СЦ | 0 | ClientRecord.service_center | clean_optional_str (обогащение карточки клиента при наличии) | — |
| D | документов всего | 2928 | SliceRecord.total | parse_strict_int; период 2026-01-01..2026-03-31, срез 2026-10-20 | — |
| E | Всего не отправлено документов | 2928 | SliceRecord.unsent | parse_strict_int; период 2026-01-01..2026-03-31, срез 2026-10-20 | — |

## Срез 20.10.26_Дубна.xlsx — 2 кв 2026

Строка заголовков: 2. Колонок: 5.

| Колонка | Заголовок Excel | Непустых ячеек | Основное поле | Преобразование | Причина исключения |
| --- | --- | ---: | --- | --- | --- |
| A | Уникальны обезличинный идентификатор | 3021 | SliceRecord.source_client_id | Сопоставление с белым списком; неизвестные -> UNKNOWN_CLIENT_ID (blocked) | — |
| B | код СЦ | 0 | ClientRecord.service_center_code | clean_optional_str; код СЦ отдельно от названия service_center; происхождение B | — |
| C | название СЦ | 0 | ClientRecord.service_center | clean_optional_str (обогащение карточки клиента при наличии) | — |
| D | документов всего | 3021 | SliceRecord.total | parse_strict_int; период 2026-04-01..2026-06-30, срез 2026-10-20 | — |
| E | Всего не отправлено документов | 3021 | SliceRecord.unsent | parse_strict_int; период 2026-04-01..2026-06-30, срез 2026-10-20 | — |

## Срез 20.10.26_Дубна.xlsx — 3 кв 2026

Строка заголовков: 2. Колонок: 5.

| Колонка | Заголовок Excel | Непустых ячеек | Основное поле | Преобразование | Причина исключения |
| --- | --- | ---: | --- | --- | --- |
| A | Уникальны обезличинный идентификатор | 3404 | SliceRecord.source_client_id | Сопоставление с белым списком; неизвестные -> UNKNOWN_CLIENT_ID (blocked) | — |
| B | код СЦ | 0 | ClientRecord.service_center_code | clean_optional_str; код СЦ отдельно от названия service_center; происхождение B | — |
| C | название СЦ | 0 | ClientRecord.service_center | clean_optional_str (обогащение карточки клиента при наличии) | — |
| D | документов всего | 3404 | SliceRecord.total | parse_strict_int; период 2026-07-01..2026-09-30, срез 2026-10-20 | — |
| E | Всего не отправлено документов | 3404 | SliceRecord.unsent | parse_strict_int; период 2026-07-01..2026-09-30, срез 2026-10-20 | — |

## База должники III кв 2025 Дубна.xlsx — 1-3 кв 25

Строка заголовков: 1. Колонок: 60.

| Колонка | Заголовок Excel | Непустых ячеек | Основное поле | Преобразование | Причина исключения |
| --- | --- | ---: | --- | --- | --- |
| A | Уникальны обезличинный идентификатор | 295 | source_client_id | str.strip(); сверка с реестром клиентов | — |
| B | Код КЦР | 0 | ClientRecord.kcr_code | clean_optional_str (обогащение карточки клиента) | — |
| C | название СЦ | 0 | ClientRecord.service_center | clean_optional_str (обогащение карточки клиента) | — |
| D | Дата создания КЦР | 0 | ClientRecord.creation_date / ClientRecord.raw_creation_date | parse_iso_date: календарная дата YYYY-MM-DD или null; исходное значение отдельно; INVALID_CREATION_DATE при ошибке | — |
| E | Время от Екб | 0 | ClientRecord.time_zone_offset | clean_optional_str (обогащение карточки клиента) | — |
| F | ФИО главного админа | 0 | ClientRecord.admin_name | clean_optional_str (истинное ФИО главного администратора) | — |
| G | Телефон админа | 0 | ClientRecord.admin_phone | clean_optional_str (телефон администратора из выделенной колонки) | — |
| H | Email админа | 0 | ClientRecord.admin_email | clean_optional_str (email администратора из выделенной колонки) | — |
| I | % долга за весь 2025 на 27.03.26 | 295 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Не включено в текущие ClientRecord/SliceRecord/HistoryRecord; сохранено без трактовки в source_preserved_values.jsonl. Нормализация требует согласования. |
| J | коммаентарий для претензии | 52 | HistoryRecord(claim_comment).text | Дословное сохранение текста претензии (52 записи), source_column='J' | — |
| K | Статус на 26.05.26 | 184 | HistoryRecord(historical_status).raw_status | occurred_at='2026-05-26', source_column='K' (184 записи) | — |
| L | Статус на 27.03.26 | 295 | HistoryRecord(historical_status).raw_status | occurred_at='2026-03-27', source_column='L' (295 записей) | — |
| M | Статус на 13.03.26 | 295 | HistoryRecord(historical_status).raw_status | occurred_at='2026-03-13', source_column='M' (295 записей) | — |
| N | Статус на 02.03.26 | 32 | HistoryRecord(historical_status).raw_status | occurred_at='2026-03-02', source_column='N' (32 записи) | — |
| O | Статус на 16.01.26 | 294 | HistoryRecord(historical_status).raw_status | occurred_at='2026-01-16', source_column='O' (294 записи) | — |
| P | Статус на 29.12.25 | 295 | HistoryRecord(historical_status).raw_status | occurred_at='2025-12-29', source_column='P' (295 записей) | — |
| Q | Статус на 5.12.25 | 295 | HistoryRecord(historical_status).raw_status | occurred_at='2025-12-05', source_column='Q' (295 записей) | — |
| R | админ в нескольких кцр | 295 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Не включено в текущие ClientRecord/SliceRecord/HistoryRecord; сохранено без трактовки в source_preserved_values.jsonl. Нормализация требует согласования. |
| S | Документов всего за 2025 год | 296 | SliceRecord.total | parse_strict_int; базовый total для 6 ретроспективных срезов 2025 года | — |
| T | Долг на 27.03.26 | 296 | SliceRecord.unsent | parse_strict_int; период 2025-01-01..2025-12-31, срез 2026-03-27 | — |
| U | % Долга на 27.03.26 | 296 | SliceRecord.raw_percentage | parse_strict_float; срез 2026-03-27 | — |
| V | Уменьшение долга октябрь vs март | 295 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Не включено в текущие ClientRecord/SliceRecord/HistoryRecord; сохранено без трактовки в source_preserved_values.jsonl. Нормализация требует согласования. |
| W | долг на 16.01.25 | 296 | SliceRecord.unsent (rejected) | Изолируется в rejected_slices.jsonl из-за противоречия даты с колонкой X (16.01.25 vs 16.01.26) | — |
| X | % Долга на 16.01.26 | 296 | SliceRecord.raw_percentage (rejected) | Изолируется в rejected_slices.jsonl (AMBIGUOUS_AS_OF_DATE:2025-01-16_vs_2026-01-16) | — |
| Y | Уменьшение долга ноябрь vs январь | 296 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Не включено в текущие ClientRecord/SliceRecord/HistoryRecord; сохранено без трактовки в source_preserved_values.jsonl. Нормализация требует согласования. |
| Z | % уменьшение долга ноябрь vs январь | 295 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Не включено в текущие ClientRecord/SliceRecord/HistoryRecord; сохранено без трактовки в source_preserved_values.jsonl. Нормализация требует согласования. |
| AA | долг на 05.12.25 | 296 | SliceRecord.unsent | parse_strict_int; период 2025-01-01..2025-12-31, срез 2025-12-05 | — |
| AB | % Долга на 05.12.25 | 296 | SliceRecord.raw_percentage | parse_strict_float; срез 2025-12-05 | — |
| AC | Уменьшение долга ноябрь vs декабрь | 296 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Не включено в текущие ClientRecord/SliceRecord/HistoryRecord; сохранено без трактовки в source_preserved_values.jsonl. Нормализация требует согласования. |
| AD | % уменьшение долга ноябрь vs декабрь | 296 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Не включено в текущие ClientRecord/SliceRecord/HistoryRecord; сохранено без трактовки в source_preserved_values.jsonl. Нормализация требует согласования. |
| AE | долг на 05.11.25 | 296 | SliceRecord.unsent | parse_strict_int; период 2025-01-01..2025-12-31, срез 2025-11-05 | — |
| AF | % Долга на 05.11.25 | 296 | SliceRecord.raw_percentage | parse_strict_float; срез 2025-11-05 | — |
| AG | долг на 31.10.25 | 295 | SliceRecord.unsent | parse_strict_int; период 2025-01-01..2025-12-31, срез 2025-10-31 (#N/A -> rejected) | — |
| AH | % Долга на 31.10.25 | 295 | SliceRecord.raw_percentage | parse_strict_float; срез 2025-10-31 (#N/A -> rejected) | — |
| AI | долг на 20.10.25 | 295 | SliceRecord.unsent | parse_strict_int; период 2025-01-01..2025-12-31, срез 2025-10-20 | — |
| AJ | отклонено | 295 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Не включено в текущие ClientRecord/SliceRecord/HistoryRecord; сохранено без трактовки в source_preserved_values.jsonl. Нормализация требует согласования. |
| AK | на исправлении | 295 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Не включено в текущие ClientRecord/SliceRecord/HistoryRecord; сохранено без трактовки в source_preserved_values.jsonl. Нормализация требует согласования. |
| AL | % долга | 295 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Не включено в текущие ClientRecord/SliceRecord/HistoryRecord; сохранено без трактовки в source_preserved_values.jsonl. Нормализация требует согласования. |
| AM | отправлено | 295 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Не включено в текущие ClientRecord/SliceRecord/HistoryRecord; сохранено без трактовки в source_preserved_values.jsonl. Нормализация требует согласования. |
| AN | принято | 295 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Не включено в текущие ClientRecord/SliceRecord/HistoryRecord; сохранено без трактовки в source_preserved_values.jsonl. Нормализация требует согласования. |
| AO | Статус на 20.10.25 | 295 | HistoryRecord(historical_status).raw_status | occurred_at='2025-10-20', source_column='AO' (295 записей) | — |
| AP | Статус на 05.11.25 | 295 | HistoryRecord(historical_status).raw_status | occurred_at='2025-11-05', source_column='AP' (295 записей) | — |
| AQ | Арина 14.10.25 | 295 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Не включено в текущие ClientRecord/SliceRecord/HistoryRecord; сохранено без трактовки в source_preserved_values.jsonl. Нормализация требует согласования. |
| AR | Отправляли посылку без описи | 295 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Не включено в текущие ClientRecord/SliceRecord/HistoryRecord; сохранено без трактовки в source_preserved_values.jsonl. Нормализация требует согласования. |
| AS | Ответственный ООВС | 0 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Нет непустых значений ниже заголовка в предоставленном файле. |
| AT | Комментарии ООВС | 295 | HistoryRecord(first_contact).text | Дословный текст комментария первого касания (source_column='AT-AX') | — |
| AU | Статус звонка | 295 | HistoryRecord(first_contact).raw_status | Дословный статус первого звонка (source_column='AT-AX') | — |
| AV | Дата перезвона | 0 | HistoryRecord(first_contact).next_contact_date / raw_next_contact | parse_iso_date + сохранение некалендарного текста и NON_DATE_NEXT_CONTACT | — |
| AW | Срок отправки от клиента | 282 | HistoryRecord(first_contact).raw_send_deadline | Сохраняется в raw_send_deadline и в объединённом raw_deadline | — |
| AX | Дата от клиента | 62 | HistoryRecord(first_contact).raw_client_date | Сохраняется в raw_client_date (62 записи) и в объединённом raw_deadline | — |
| AY | Звоним в январе? | 295 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Не включено в текущие ClientRecord/SliceRecord/HistoryRecord; сохранено без трактовки в source_preserved_values.jsonl. Нормализация требует согласования. |
| AZ | Ответственный ООВС | 0 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Нет непустых значений ниже заголовка в предоставленном файле. |
| BA | Комментарий к повторному звонку | 152 | HistoryRecord(second_contact).text | Дословный комментарий повторного звонка в январе (source_column='BA-BE') | — |
| BB | Статус звонка в январе | 278 | HistoryRecord(second_contact).raw_status | Дословный статус повторного звонка в январе (278 записей) | — |
| BC | Дата перезвона | 5 | HistoryRecord(second_contact).next_contact_date / raw_next_contact | parse_iso_date + сохранение сырого текста и NON_DATE_NEXT_CONTACT | — |
| BD | Срок отправки по клиенту | 229 | HistoryRecord(second_contact).raw_send_deadline | Сохраняется в raw_send_deadline (229 записей) и в raw_deadline | — |
| BE | Срок от клиента | 19 | HistoryRecord(second_contact).raw_client_date | Сохраняется в raw_client_date (19 записей) и в raw_deadline | — |
| BF | Инфа по документам от клиента | 1 | HistoryRecord(incident_or_expert_request).raw_deadline / text | Сохраняется в истории инцидентов/экспертных запросов (source_column='BF-BH') | — |
| BG | № Инцидента по отсрочке | 1 | HistoryRecord(incident_or_expert_request).text | Сохраняется в истории инцидентов/экспертных запросов (source_column='BF-BH') | — |
| BH | Запрос эксперта | 6 | HistoryRecord(incident_or_expert_request).raw_status / text | Сохраняется в истории инцидентов/экспертных запросов (source_column='BF-BH') | — |

## База I кв 2026 Дубна.xlsx — Все КЦР

Строка заголовков: 2. Колонок: 42.

| Колонка | Заголовок Excel | Непустых ячеек | Основное поле | Преобразование | Причина исключения |
| --- | --- | ---: | --- | --- | --- |
| A | Уникальны обезличинный идентификатор | 115 | source_client_id | str.strip(); сверка с реестром клиентов | — |
| B | Код КЦР | 0 | ClientRecord.kcr_code | clean_optional_str (обогащение карточки клиента) | — |
| C | (без заголовка) | 0 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Нет непустых значений ниже заголовка в предоставленном файле. |
| D | Название | 0 | ClientRecord.organization_name | clean_optional_str (обогащение карточки клиента) | — |
| E | СЦ | 0 | ClientRecord.service_center | clean_optional_str (обогащение карточки клиента) | — |
| F | (без заголовка) | 0 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Нет непустых значений ниже заголовка в предоставленном файле. |
| G | (без заголовка) | 0 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Нет непустых значений ниже заголовка в предоставленном файле. |
| H | (без заголовка) | 0 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Нет непустых значений ниже заголовка в предоставленном файле. |
| I | (без заголовка) | 0 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Нет непустых значений ниже заголовка в предоставленном файле. |
| J | Время от Екб | 115 | ClientRecord.time_zone_offset | clean_optional_str (обогащение 115 карточек клиентов с фиксацией attribute_provenance) | — |
| K | ФИО главного админа | 0 | ClientRecord.admin_name | clean_optional_str (ФИО главного администратора, отдельно от manager_name) | — |
| L | Телефон админа | 0 | ClientRecord.admin_phone | clean_optional_str (телефон администратора из выделенного столбца) | — |
| M | Email админа | 0 | ClientRecord.admin_email | clean_optional_str (email администратора из выделенного столбца) | — |
| N | Спич | 115 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Не включено в текущие ClientRecord/SliceRecord/HistoryRecord; сохранено без трактовки в source_preserved_values.jsonl. Нормализация требует согласования. |
| O | стаус на 21.05.26 | 115 | HistoryRecord(historical_status).raw_status | occurred_at='2026-05-21', source_column='O' (115 записей) | — |
| P | ИНН | 0 | ClientRecord.inn | clean_optional_str (обогащение карточки клиента) | — |
| Q | КПП | 0 | ClientRecord.kpp | clean_optional_str (обогащение карточки клиента) | — |
| R | Действует ли КЦР | 115 | ClientRecord.is_active | clean_optional_str (обогащение карточки клиента) | — |
| S | Блокировка в сервисе | 115 | ClientRecord.is_blocked | clean_optional_str (обогащение карточки клиента) | — |
| T | ФИО менеджера | 0 | ClientRecord.manager_name | clean_optional_str (сохраняется строго в manager_name) | — |
| U | ФИО менеджера из КП | 0 | ClientRecord.manager_name | Резервное поле manager_name, если T пусто | — |
| V | Всего документов весь 2025 | 115 | SliceRecord.total | parse_strict_int; период 2025-01-01..2025-12-31, срез 2026-05-21 | — |
| W | Долг весь 2025 на 21.05.26 | 115 | SliceRecord.unsent | parse_strict_int; период 2025-01-01..2025-12-31, срез 2026-05-21 | — |
| X | % Долга весь 2025 на 21.05.26 | 115 | SliceRecord.raw_percentage | parse_strict_float; срез 2026-05-21 | — |
| Y | Всего документов 2-4 кв 2025 | 114 | SliceRecord.total | parse_strict_int; период 2025-04-01..2025-12-31, срез 2026-05-21 | — |
| Z | Долг2-4 кв 2025 на 21.05.26 | 114 | SliceRecord.unsent | parse_strict_int; период 2025-04-01..2025-12-31, срез 2026-05-21 | — |
| AA | % Долга 2-4кв 2025 на 21.05.26 | 114 | SliceRecord.raw_percentage | parse_strict_float; срез 2026-05-21 | — |
| AB | Всего документов 1 кв 2026 | 107 | SliceRecord.total | parse_strict_int; период 2026-01-01..2026-03-31, срез 2026-05-21 | — |
| AC | Долг 1 кв 2026 на 21.05.26 | 107 | SliceRecord.unsent | parse_strict_int; период 2026-01-01..2026-03-31, срез 2026-05-21 | — |
| AD | % Долга 1 кв 2026 на 21.05.26 | 107 | SliceRecord.raw_percentage | parse_strict_float; срез 2026-05-21 | — |
| AE | Всего документов 2 кв 2025-1 кв 2026 на 21.05.26 | 115 | SliceRecord.total | parse_strict_int; период 2025-04-01..2026-03-31, срез 2026-05-21 | — |
| AF | Долг 2 кв 2025-1 кв 2026 на 21.05.26 | 115 | SliceRecord.unsent | parse_strict_int; период 2025-04-01..2026-03-31, срез 2026-05-21 | — |
| AG | % Долга 2 кв 2025-1 кв 2026 на 21.05.26 | 113 | SliceRecord.raw_percentage | parse_strict_float; срез 2026-05-21 | — |
| AH | Ответственный ООВС | 0 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Нет непустых значений ниже заголовка в предоставленном файле. |
| AI | Комментарии ООВС | 115 | HistoryRecord(operator_contact).text | Дословный комментарий оператора I кв. 2026 (115 записей, source_column='AI-AM') | — |
| AJ | Статус звонка | 115 | HistoryRecord(operator_contact).raw_status | Дословный статус звонка I кв. 2026 (115 записей, source_column='AI-AM') | — |
| AK | Дата перезвона | 0 | HistoryRecord(operator_contact).next_contact_date / raw_next_contact | parse_iso_date + сохранение некалендарного текста и NON_DATE_NEXT_CONTACT | — |
| AL | Срок отправки от клиента | 93 | HistoryRecord(operator_contact).raw_send_deadline | Сохраняется в raw_send_deadline (93 записи) и в raw_deadline | — |
| AM | Дата от клиента | 0 | HistoryRecord(operator_contact).raw_client_date | Сохраняется в raw_client_date и в raw_deadline | — |
| AN | № Инцидента по отсрочке | 0 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Нет непустых значений ниже заголовка в предоставленном файле. |
| AO | Запрос эксперта | 0 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Нет непустых значений ниже заголовка в предоставленном файле. |
| AP | (без заголовка) | 0 | — | Сохранение исходного значения и координат в source_preserved_values.jsonl при наличии значения. | Нет непустых значений ниже заголовка в предоставленном файле. |

