# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

Документация и общение с пользователем — на русском языке.

## Что это

Синхронизация `БД инвентаризации -> Active Directory` (основное направление) с обратным push'ем
нескольких полей `AD -> инвентаризация` (Login, Email). Один рабочий скрипт:
[inventory-to-ad.ps1](inventory-to-ad.ps1). Чисто обратное направление (`AD -> Inventory`) живёт в
отдельном репозитории `../ad-inventory-sync` — не путать по имени, оно почти такое же.
Внешний скрипт увольнения берётся из `../ad-usermanagement` (spo0okie/ad-usermanagement).

## Запуск

Windows PowerShell 5.1 + модуль `ActiveDirectory` (RSAT), запуск под учёткой с правами записи в AD.

Полный прогон по всем OU из конфига:

```bash
inventory-to-ad.cmd
```

Один пользователь (основной способ отладки — обрабатывается только он, OU игнорируются):

```bash
powershell.exe -noprofile -executionpolicy bypass -file inventory-to-ad.ps1 ivanov
```

Автотестов нет. Штатный цикл проверки изменений: выставить `$write_ad=$false` и
`$write_inventory=$false` в `../config.priv.ps1` (режим RO — все записи заменяются строчками
«AD RO mode» / «INV: RO mode»), прогнать одного пользователя, посмотреть diff в логе, затем
включать запись.

## Внешние зависимости и раскладка на диске

Скрипт делает dot-source по путям `$PSScriptRoot\..\`, поэтому репозиторий обязан лежать
подпапкой рядом с конфигом и библиотеками:

```
<рабочая папка>/
  config.priv.ps1          # НЕ в репозитории: креды, OUDN, флаги, путь к логу
  libs.ps1/                # отдельный репозиторий spo0okie/ps1.libs
    lib_funcs.ps1          # spooLog / errorLog / warningLog / debugLog, correctMobile, correctPhonesList
    lib_inventory.ps1      # REST-клиент к инвентаризации (Basic-auth)
    lib_usr_ad.ps1         # DisableADUser, PrepareOU, CreateADUser
  inventory-ad-sync/       # этот репозиторий
```

Ключевые функции библиотек:
- `getInventoryObj <model> <name> <@{доп. параметры}>` — `GET /<model>/search`, возвращает объект
  или `$false` (404 тоже `$false`; шум в логе глушится `$global:skip404errors`).
- `callInventoryRestMethod <method> <model> <action> <@{параметры}> <raw>` — произвольный вызов,
  при `raw=$true` возвращает тело ответа строкой (так тянется внутренний номер телефона).
- `setInventoryData 'users' <хэш с id>` — `PUT /users/<id>`; без валидного `id` сделает `POST`
  (создание), чего в этом скрипте не хотят никогда — `pushUserData` всегда передаёт `id`.
- `correctMobile` приводит номер к `+7(987)654-3210`, `correctPhonesList` делает то же для списка
  через запятую и обрезает результат по лимиту поля `mobile` (64 символа).
- `DisableADUser` (lib_usr_ad.ps1) сбрасывает пароль, отключает учётку и переносит её из `$u_OUDN`
  в `$f_OUDN`, создавая недостающие OU.

### Конфиг

Из `config.priv.ps1` скрипт использует: `$inventory2ad_sync`, `$write_ad`, `$auto_dismiss`,
`$auto_dismiss_exclude`, `$dismiss_script`, `$multiorg_support`, `$exchange_domains`,
`$inventory_dateformat`, `$write_inventory`, `$logfile`, `$inventory_RESTapi_URL`,
`$inventory_user_login` / `$inventory_user_password`.

Массовый режим ходит по `$inventory2ad_sync` — это **массив хэшей**, по одному на OU:

```powershell
$inventory2ad_sync=@(
    @{u_OUDN="OU=Пользователи,DC=domain,DC=local"; f_OUDN="OU=Уволенные,DC=domain,DC=local"}
)
```

Полный пример конфига — в [README.md](README.md). Если `$inventory2ad_sync` не определён, `foreach`
просто не сделает ни одной итерации и скрипт молча завершится, ничего не синхронизировав, —
первое, что стоит проверять при «синхронизация ничего не делает».

## Логика синхронизации

Точка входа — в конце файла: либо один пользователь из `$args[0]`, либо
`Get-ADUser -Filter {enabled -eq $true} -SearchBase $u_OUDN` по каждому элементу
`$inventory2ad_sync`. Каждый объект уходит в `ParseUser`.

`FindUser` ищет пользователя в инвентаризации строго по очереди, первое совпадение выигрывает:

1. `login` = `sAMAccountName` (уникален),
2. `num` + `org` = `employeeID` + `employeeNumber` (пара уникальна),
3. `uid` = `adminDescription` (ИНН/СНИЛС; может найтись несколько — приоритет расставляет сама
   инвентаризация),
4. `name` = `displayName`,
5. `num` = `displayName` (для сквозной нумерации табельных — табельный вместо ФИО).

Все запросы идут с `expand=ln,mn,fn,orgStruct,org`. Не нашли — `warningLog` и возврат строки
`'error'`, которую `ParseUser` сравнивает как `$invUser -eq "error"` и пропускает пользователя.

Обработка увольнения (самая нетривиальная часть):
- Если найденная запись помечена `Uvolen=1`, но `resign_date` **ещё не наступила** — увольнение
  отменяется прямо в объекте (`$invUser.Uvolen = 0`) и пишется предупреждение.
- Если дата уже прошла — ищутся другие трудоустройства: по `uid` (из AD `adminDescription`, иначе
  из инвентаризации), а при отсутствии `uid` — по ФИО (хуже: не отсекает однофамильцев).
- Реальное увольнение в `ParseUser` делается только при `$auto_dismiss`, при этом
  `$auto_dismiss_exclude` проверяется всегда (даже с выключенным `$auto_dismiss` — тогда просто
  пишется «Deactivation needed!»). Если задан `$dismiss_script`, вызывается он с логином в
  аргументе, иначе — `DisableADUser`. После увольнения `ParseUser` сразу выходит.

Соответствие полей (имена в инвентаризации — транслитом):

| AD | Inventory | направление / примечание |
|---|---|---|
| `name`, `cn` | `Ename` | переименование объекта через `Rename-AdObject` |
| `displayName` | `Ename` | |
| `givenName` | `fn` + `mn` | |
| `sn` | `ln` | |
| `department` | `orgStruct.name` | обрезается до 64 символов |
| `company` | `org.uname` | |
| `title` | `Doljnost` | обрезается до 128 символов |
| `adminDescription` | `uid` | |
| `employeeID` | `employee_id` | табельный номер |
| `employeeNumber` | `org_id` | только при `$multiorg_support` |
| `mobile` | `Mobile` | через `correctPhonesList` |
| `pager` | `GET /phones/search-by-user` (raw) | внутренний номер, отдельный запрос |
| `telephoneNumber` | — | из инвентаризации не берётся, только нормализуется формат |
| `mail` | `Email` | направление зависит от домена: см. ниже |
| `sAMAccountName` | `Login` | **AD -> инвентаризация** |

Почта: если адрес в AD оканчивается на один из `$exchange_domains`, источником считается AD и
значение пушится в инвентаризацию (`pushUserData`); иначе наоборот — в AD пишется `Email` из
инвентаризации.

Запись в AD собирается в объекте `$user` и уходит одним `Set-AdUser -Instance $user` в конце
(флаг `$needUpdate`), переименование — отдельно (флаг `$needRename`).

## Соглашения кода и грабли

- **Кодировка**: `.ps1` этого репозитория — UTF-8 **без BOM**, CRLF (коммит «saved in unicode»);
  русский текст есть только в комментариях, все строковые литералы и логи — ASCII/английский.
  Так и держать: не тащить кириллицу в строки и не переводить файл в другую кодировку.
  Библиотеки в `../libs.ps1/` при этом разной кодировки (`lib_usr_ad.ps1` — Windows-1251),
  их тоже не перегонять.
- Очистка поля в AD **не** делается через `Set-AdUser -Instance` — пустое значение так не
  запишется. Для этого в коде отдельные ветки `Get-ADUser ... | Set-AdUser -Clear <attr>`
  (`mobile`, `pager`, `sn`, `givenName`). Новое обнуляемое поле надо оформлять так же.
- Новый атрибут для синхронизации нужно добавить в **оба** вызова `Get-ADUser -properties ...`
  (ветка одного пользователя и ветка обхода OU), иначе поле молча приедет пустым.
- Сравнения делаются через `[string]$a -ne [string]$b` — осознанно: `$null` и пустая строка
  должны считаться равными.
- Любая запись в AD обязана быть под `if ($write_AD)` с веткой-заглушкой `spooLog(... RO mode)`,
  любая запись в инвентаризацию — только через `pushUserData` (там проверка `$write_inventory`).
- Логирование только через `spooLog` / `warningLog` / `errorLog` / `debugLog` — они пишут и в
  консоль, и в `$global:logfile`. `user.log` в `.gitignore`, коммитить его не надо.
- Проверка «поле не пустое» — всегда `.Length -gt 0`, а не `-gt 0`: у строки сравнение с числом
  даёт строковое сравнение с `"0"` и работает случайно.
- Ограничения схемы AD, о которых напоминает шапка скрипта: `mobile` — 64, `title` — 128,
  `department` — 64. Проверить лимит атрибута:
  `dsquery * "cn=Schema,cn=Configuration,dc=Domain,dc=local" -Filter "(LDAPDisplayName=mobile)" -attr rangeUpper`.
- `_arch/` — старая версия скрипта (`SAPsync_users_v2.ps1`), в `.gitignore`, как образец не брать.
