# AD-Inventory-sync
Скрипт двусторонней синхронизации AD с БД инвентаризации

При синхронизации ищет переданного пользователя в БД сначала по табельному номеру, затем по ФИО  
Для увольнения сотрудников вызывается скрипт из скриптов управления https://github.com/spo0okie/ad-usermanagement

## Синхронизирует поля БД -> AD:
* ФИО (Название объекта АД, Выводимое имя, Имя, Фамилия)
* Табельный номер и ИД организации
* Организация
* Подразделение
* Должность
* Мобильный, внутренний, городской номер телефона
* UID

## Синхронизирует поля БД <- AD:
* Login
* E-Mail (в зависимости от домена)
* Городской номер телефона

Инклудит файл ..\config.priv.ps1 следующего содержания
```powershell
#домен организации
$domain="domain.local"
#подразделение где хранятся пользователи
$u_OUDN="OU=Пользователи,DC=domain,DC=local"
#подразделение куда увольняются пользователи
$f_OUDN="OU=Уволенные,DC=domain,DC=local"
#писать ли изменения в АД?
$write_ad=$false
#разрешить автоматическое увольнение при синхронизации
$auto_dismiss=$false
#исключить автоматическое уольнение этих логинов:
$auto_dismiss_exclude=@(
	'pupkin.v',
	'lohankin.v', 
	'lenin.v'
)

#URL сервиса запросов в таблицу пользователей САП
$inventory_RESTapi_URL="http://inventory.domain.local/web/api"
#поддержка нескольких организаций
$multiorg_support=$false
#формат даты (приема, увольнения, др) в САП
$dateformat_SAP='yyyy-MM-dd'
#писать ли изменения в БД инвентаризации
$write_inventory=$false

#логфайл синхронизации
$sync_logfile="C:\Joker\Works\PS\user_management\SAPsync\log\user_sync.log"
```


## История изменений
v5.0 + Перевод на авторизацию REST API  
v4.0 + Использование UID(ИНН.СНИЛС и т.п.) для идентификации сотрудника  
v3.2 + Возможность синхронизации одного пользователя  
v3.1 + Исключения сотрудников из автоувольнения  
v3.0 + Режим только чтения для АД  
     + Режим только чтение для БД инвентаризации  
     + Опции для выбора источника телефонного номера  
     + Опции для отключения автоувольнения  
     * Небольшая шлифовка при сопряжении с САП.  
v2.3 + Поддержка нескольких организаций  
v2.2 + Учет даты увольния при увольнении  
v2.1 + Обработка нескольких телефонов через запятую, ограничение длины полей  
v2.0 + Поддержка нескольких организаций  
      Убрана обертка вокруг БД Инвентаризации для взаимодействия с этими скриптами  
      Теперь все взаимодействие ведется через каноничный REST API  
v1.7 * конфиг вынесен во внешний файл  
v1.6 + Добавлено увольнение сотрудника при увольнении в БД через внешний скрипт управления  
v1.5 + Городской номер приводится к тому же формату, что и мобильный  
v1.4 + Синхронизация почты, внутреннего номера телефона, мобильного номера телефона  
       в сторону САП  
     + мобильные телефонные номера сначала приводятся к формату +7(ХХХ)ХХХ-ХХХХ  
v1.3 Хранение табельного номера переведено на EmployeeID  
v1.2 Улучшена корректировака имен. В т.ч. переименование самого объекта  
v1.1 Улучшена корректировка имени.  
     Пользователя можно искать указав в качестве имени табельный номер  
v1.0 Initial commit  

