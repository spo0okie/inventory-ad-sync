#Скрипт управление пользователем в АД:
#
#в качестве параметра принимает JSON объект по этому пользователю
#
#ищет переданного пользователя в АД сначала по табельному номеру,
#затем по ФИО
#
#Синхронизирует поля:
# - ФИО
# - Табельный номер
# - Должность
# - Подразделение
# - Организация
# 

#как посмотреть лимит на длину поля? например для mobile вот так:
#dsquery * "cn=Schema,cn=Configuration,dc=yamalgazprom,dc=local" -Filter "(LDAPDisplayName=mobile)" -attr rangeUpper

#у нас были затыки с полями
#mobile (64)
#title (128)
#department (64)

#Windows PowerShell 5.1 по умолчанию предлагает серверу только SSL3/TLS1.0 -
#без TLS1.2 все REST-запросы к инвентаризации падают с ошибкой создания защищенного канала
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

. "$($PSScriptRoot)\..\config.priv.ps1"
. "$($PSScriptRoot)\..\libs.ps1\lib_funcs.ps1"
. "$($PSScriptRoot)\..\libs.ps1\lib_inventory.ps1"
. "$($PSScriptRoot)\..\libs.ps1\lib_usr_ad.ps1"

#связи, которые нужно подтягивать вместе с кадровой записью
$inventory_user_expand='ln,mn,fn,orgStruct,org'

#приоритетная организация по умолчанию (может переопределяться поэлементно в $inventory2ad_sync)
$stickyOrgDefault=$stickyOrg

#кадровая запись, с которой учетка была перепривязана на этом прогоне ($false - перепривязки не было)
$global:userRebindFrom=$false


#запись данных о пользователе в БД
function pushUserData() {
	param
	(
		[string]$id,
		[string]$field,
		[string]$value
	)
	$params = @{
		$field=$value;
		id=$id;
	}

	if ($write_inventory) {
		setInventoryData 'users' $params
	} else {
		spooLog("invPush: skip user #$id $field = $value INV: RO mode")
	}
}

#разбирает дату из инвентаризации в [datetime]
#возвращает $null если даты нет или она не в формате $inventory_dateformat
function parseInventoryDate() {
	param
	(
		[string]$value
	)
	if ($value.Length -eq 0) {return $null}
	try {
		return [datetime]::parseexact($value, $inventory_dateformat, $null)
	} catch {
		debugLog("cant parse date [$value] with format [$inventory_dateformat]")
		return $null
	}
}

#действующее ли трудоустройство:
#либо не помечено уволенным, либо дата увольнения еще не наступила
function isActiveEmployment() {
	param
	(
		[object]$employment
	)
	if ($employment.Uvolen -ne "1") {return $true}
	$resign_date=parseInventoryDate $employment.resign_date
	#уволен без даты увольнения - считаем что уже уволен
	if ($null -eq $resign_date) {return $false}
	return ((Get-Date) -lt $resign_date)
}

#относится ли трудоустройство к указанной организации
#организацию можно задавать как ID, так и именем (юр. названием или брендом)
function isOrgEmployment() {
	param
	(
		[object]$employment,
		$org
	)
	if ( -not $org) {return $false}
	if ("$org" -match '^\d+$') {
		return ([string]$employment.org_id -eq [string]$org)
	}
	return (
		([string]$employment.org.uname -eq [string]$org) -or
		([string]$employment.org.name -eq [string]$org)
	)
}

#относится ли трудоустройство к приоритетной организации ($stickyOrg)
function isStickyOrgEmployment() {
	param
	(
		[object]$employment
	)
	return (isOrgEmployment $employment $stickyOrg)
}

#приводит $affiliatedOrgs к списку групп
#PowerShell разворачивает @( @(1,2) ) обратно в плоский @(1,2), поэтому единственная группа
#приезжает сюда просто списком организаций - собираем такие одиночные элементы в одну группу
#(иначе каждая организация оказалась бы сама по себе и перепривязка встала бы совсем)
function NormalizeOrgGroups() {
	param
	(
		$groups
	)
	$result=@()
	$loose=@()
	foreach ($group in $groups) {
		if ($group -is [array]) {
			$result+=,@($group)
		} else {
			$loose+=$group
		}
	}
	if ($loose.Count) {$result+=,@($loose)}

	#запятая не дает развернуть единственную группу обратно в плоский список
	return ,$result
}

#группа аффилированных организаций ($affiliatedOrgs), в которую входит трудоустройство
#$false - организация не состоит ни в одной группе
function AffiliatedOrgGroup() {
	param
	(
		[object]$employment
	)
	foreach ($group in $affiliatedOrgs) {
		foreach ($org in $group) {
			if (isOrgEmployment $employment $org) {return $group}
		}
	}
	return $false
}

#допустимо ли перевести учетку с трудоустройства $from на трудоустройство $to:
#внутри одной организации - всегда, между разными - только внутри группы аффилированных
#(организация вне групп сама по себе: за ее пределы учетку не уводим)
function CanRebindEmployment() {
	param
	(
		[object]$from,
		[object]$to
	)
	if ([string]$from.org_id -eq [string]$to.org_id) {return $true}

	$group=AffiliatedOrgGroup $from
	if ( -not $group) {return $false}

	foreach ($org in $group) {
		if (isOrgEmployment $to $org) {return $true}
	}
	return $false
}

#выкидывает из списка трудоустройства, на которые учетку переводить нельзя:
#организация должна быть той же самой либо аффилированной с текущей
function SelectRebindableEmployments() {
	param
	(
		$employments,
		[object]$current
	)
	if ($current -isnot [PSCustomObject]) {return @($employments)}

	$allowed=@()
	foreach ($employment in $employments) {
		if (([string]$employment.id -eq [string]$current.id) -or (CanRebindEmployment $current $employment)) {
			$allowed+=$employment
		} else {
			debugLog("employment #$($employment.id) (org $($employment.org_id)) is not affiliated with org $($current.org_id), skipped")
		}
	}

	if ( -not $allowed.Count) {return @($employments)}

	return @($allowed)
}

#запрашивает ВСЕ кадровые записи, подходящие под фильтр (в отличие от search, отдающего одну)
function FetchEmployments() {
	param
	(
		[hashtable]$filter
	)
	$filter['expand']=$inventory_user_expand
	$filter['per-page']=100
	$employments=callInventoryRestMethod 'GET' 'users' 'filter' $filter
	#404/ошибка запроса приезжают как $false, пустой список - как $null
	if (($null -eq $employments) -or ($employments -is [bool])) {return @()}
	return @($employments)
}

#выкидывает из списка трудоустройства, занятые другой учеткой АД: у человека бывает несколько
#логинов (напр. по одному на организацию), и каждый должен держаться своего трудоустройства,
#иначе обе учетки сойдутся на одной кадровой записи и начнут отнимать ее друг у друга
#запись без логина ничья - на нее претендовать можно
function SelectOwnEmployments() {
	param
	(
		$employments,
		[string]$login
	)
	$own=@()
	foreach ($employment in $employments) {
		$employmentLogin=[string]$employment.Login
		if (($employmentLogin.Length -eq 0) -or ($employmentLogin -eq $login)) {
			$own+=$employment
		} else {
			debugLog("$($login): employment #$($employment.id) belongs to [$employmentLogin], skipped")
		}
	}

	#все записи человека заняты другими учетками - выбирать не из чего, работаем как нашли
	if ( -not $own.Count) {
		debugLog("$($login): all employments belong to other AD accounts")
		return @($employments)
	}

	return @($own)
}

#оставляет только те трудоустройства, на которые учетку вообще можно переключить:
#уволенная запись целью переключения не является - приоритетной может быть только действующая
#текущая привязка остается в списке всегда: если переключаться некуда, синхронизируемся с ней
#(и увольняем учетку), но с уволенной на уволенную не прыгаем
function SelectSwitchTargets() {
	param
	(
		$employments,
		[object]$current
	)
	$targets=@()
	foreach ($employment in $employments) {
		$isCurrent=(($current -is [PSCustomObject]) -and ([string]$employment.id -eq [string]$current.id))
		if (($employment.Uvolen -ne "1") -or $isCurrent) {
			$targets+=$employment
		} else {
			debugLog("employment #$($employment.id) is dismissed, not a switch target")
		}
	}

	#ни текущей привязки, ни действующих записей - выбираем из того, что есть
	if ( -not $targets.Count) {return @($employments)}

	return @($targets)
}

#выбирает из списка трудоустройств приоритетное:
#действующее > в приоритетной организации > позднее уволенное > лучший тип трудоустройства (Persg) >
#текущая привязка > более новая запись
#дата увольнения важна именно среди уволенных: если выбрать запись без даты, скрипт не поймет,
#что человек уже уволен, и не отключит учетку (у действующих записей этот критерий нейтрален)
#текущая привязка ($current) выигрывает при равенстве по существу: перецепляемся, только если
#новая запись реально лучше (стала действующей, приоритетная организация, лучший тип трудоустройства),
#а не просто моложе - иначе получаем бессмысленные переключения между равнозначными записями
function PickEmployment() {
	param
	(
		$employments,
		$current=$null
	)
	return @($employments | Sort-Object `
		@{Expression={if (isActiveEmployment $_) {0} else {1}}},
		@{Expression={if (isStickyOrgEmployment $_) {0} else {1}}},
		@{Expression={
			if (isActiveEmployment $_) {0} else {
				$resign_date=parseInventoryDate $_.resign_date
				if ($null -eq $resign_date) {0} else {$resign_date.Ticks}
			}
		}; Descending=$true},
		@{Expression={if ("$($_.Persg)" -match '^\d+$') {[int]$_.Persg} else {[int]::MaxValue}}},
		@{Expression={if (($current -is [PSCustomObject]) -and ([string]$_.id -eq [string]$current.id)) {0} else {1}}},
		@{Expression={[int]$_.id}; Descending=$true}
	)[0]
}

#ФИО пользователя так, как оно записано в АД
#обычно это displayName, но учетку могли завести скриптом (New-ADUser -Name ... без -DisplayName)
#и тогда имя есть только в самом объекте - иначе поиск по ФИО и по табельнику вместо ФИО
#не отработает вообще, и учетка навсегда останется "не найденной"
function ADUserName() {
	param
	(
		[object]$user
	)
	foreach ($value in @($user.displayName,$user.name,$user.cn)) {
		if (([string]$value).Length -gt 0) {return [string]$value}
	}
	return ''
}

#последовательный поиск кадровой записи, от которой отталкиваемся при поиске всех трудоустройств
#возвращает объект записи или $false
function FindAnchorEmployment() {
	param
	(
		[object]$user
	)

	$adName=ADUserName $user

	#Если у нас есть только табельный - считаем что организация=1
	$org_id=$user.employeeNumber
	if (($org_id.Length -eq 0) -or ( -not $multiorg_support)) {
		#если организация не заявлена, то первая
		#эта ситуация скорее всего возникнет при переходе от инвентаризации версии под одну организацию
		#к инвентаризации версии под множество. Когда БД уже с учетом организаций а АД еще нет
		$org_id=1
	}

	$expand=$inventory_user_expand
	#Ищем пользователя последовательно:

	#Если у нас есть Логин - ищем по Логину - должна быть конкретная запись, т.к. несколько логинов быть не должно
	$invUser=getInventoryObj 'users' '' @{
		login=$user.sAMAccountname;
		expand=$expand;
	}

	#Если у нас есть организация и табельный - ищем конкретного, т.к. комбинация табельный и орг тоже уникальная
	if (($invUser -isnot [PSCustomObject]) -and ($user.employeeID.Length -gt 0)) {$invUser=getInventoryObj 'users' '' @{
		num=$user.employeeID;
		org=$org_id;
		expand=$expand;
	}}
	
	#Если у нас есть ИНН ищем по нему - тут может найтись несколько и выбирается приоритетная сортировка в самой инвентори (лучший тип трудоустройства и не уволен)
	if (($invUser -isnot [PSCustomObject]) -and ($user.adminDescription.Length -gt 0)) {$invUser=getInventoryObj 'users' '' @{
		uid=$user.adminDescription;
		expand=$expand;
	}}

	#Если у нас есть ФИО - ищем по ФИО - тут тоже несколько и тоже лучший
	if (($invUser -isnot [PSCustomObject]) -and ($adName.Length -gt 0)) {$invUser=getInventoryObj 'users' '' @{
		name=$adName;
		expand=$expand;
	}}

	#Последний сценарий - табельник вместо ФИО: учетку завели новому сотруднику, вписав в имя
	#табельный номер, а ФИО подтянется из инвентаризации на первой же синхронизации
	#(находиться должен один, т.к. так делают только при сквозной нумерации табельных)
	if (($invUser -isnot [PSCustomObject]) -and ($adName.Length -gt 0)) {$invUser=getInventoryObj 'users' '' @{
		num=$adName;
		expand=$expand;
	}}
	
	if ($invUser -isnot [PSCustomObject]) {return $false}

	return $invUser
}

#загрузить пользователя из Инвентаризации через REST API
#у человека может быть несколько трудоустройств (в т.ч. совместительство), поэтому
#выбираем не первое попавшееся, а приоритетное - и пересматриваем выбор на каждом прогоне,
#иначе учетка залипает на записи, к которой ее привязали (напр. на совместительстве,
#подхваченном в выходные между увольнением в пятницу и приемом в понедельник)
function FindUser() {
	param
	(
		[object]$user
	)

	#запись, от которой отталкиваемся: она дает нам UID/ФИО человека
	$anchor=FindAnchorEmployment $user

	#ключ личности: UID записи, иначе UID из АД, иначе ФИО (хуже, т.к. не исключает однофамильцев)
	#@() обязательно: возврат из функции разворачивает массив из одного элемента в объект,
	#а у объекта нет .Count - и единственное трудоустройство считалось бы ненайденным
	$employments=@()
	if (($anchor -is [PSCustomObject]) -and $anchor.uid) {
		$employments=@(FetchEmployments @{uid=$anchor.uid})
	} elseif ($user.adminDescription) {
		$employments=@(FetchEmployments @{uid=$user.adminDescription})
	}

	if ( -not $employments.Count) {
		$name=ADUserName $user
		if (($anchor -is [PSCustomObject]) -and ($anchor.Ename.Length -gt 0)) {$name=$anchor.Ename}
		if ($name.Length -gt 0) {$employments=@(FetchEmployments @{name=$name})}
	}

	#подстраховка: запись, найденная по логину/табельному, могла не попасть в выборку
	#(нет UID, другое написание ФИО) - тогда работаем хотя бы по ней
	if ($anchor -is [PSCustomObject]) {
		$anchorListed=$false
		foreach ($employment in $employments) {
			if ([string]$employment.id -eq [string]$anchor.id) {$anchorListed=$true}
		}
		if ( -not $anchorListed) {$employments=@($employments)+@($anchor)}
	}

	#Если все-таки не нашли
	if ( -not $employments.Count) {
		warningLog("user ["+$user.sAMAccountname+"] with Name ["+(ADUserName $user)+"] - not found in inventory (searched by login, num, uid and name)")
		return 'error'
	}

	$employments=@(SelectOwnEmployments $employments $user.sAMAccountname)
	$employments=@(SelectRebindableEmployments $employments $anchor)
	$employments=@(SelectSwitchTargets $employments $anchor)

	$invUser=PickEmployment $employments $anchor

	if ($employments.Count -gt 1) {
		debugLog("$($user.sAMAccountname): $($employments.Count) employments found, using #$($invUser.id) (org $($invUser.org_id), num $($invUser.employee_id), Persg $($invUser.Persg))")
	}

	#переключение учетки на другое трудоустройство - событие заметное, пишем в лог
	$global:userRebindFrom=$false
	if (($anchor -is [PSCustomObject]) -and ([string]$anchor.id -ne [string]$invUser.id)) {
		spooLog($user.sAMAccountname+": re-binding from employment #"+$anchor.id+" (org "+$anchor.org_id+", num "+$anchor.employee_id+") to #"+$invUser.id+" (org "+$invUser.org_id+", num "+$invUser.employee_id+")")
		$global:userRebindFrom=$anchor
	}

	#предупреждаем о предстоящем увольнении, если дата еще не наступила
	if (($invUser.Uvolen -eq "1") -and (isActiveEmployment $invUser)) {
		warningLog("user ["+$user.sAMAccountname+"] with Name ["+$user.displayName+"] - to be dismissed @ "+$invUser.resign_date)
	}

	return $invUser
}

#переносит связи и атрибуты (внутренний телефон, техника, ПК, лицензии, доступы, журнал входов)
#с прежней кадровой записи на ту, к которой теперь привязана учетка: иначе все нажитое остается
#висеть на старом табельнике
#статус источника не смотрим: нажитое принадлежит человеку и должно следовать за учеткой.
#если переносить только с уволенного, то в сценарии "уволен в пятницу - принят в понедельник"
#данные уедут в выходные на совместительство и обратно в приоритетную организацию уже не вернутся
#(в самой инвентаризации логика та же: все забирает та запись, на которой оказался логин)
function MigrateEmployment() {
	param
	(
		[object]$source,
		[object]$destination
	)
	if ($source -isnot [PSCustomObject]) {return}
	if ([string]$source.id -eq [string]$destination.id) {return}

	if ( -not $write_inventory) {
		spooLog("invPush: skip migrate employment #$($source.id) -> #$($destination.id) INV: RO mode")
		return
	}

	#параметры именно в query string: Yii подставляет в аргументы действия только их, но не тело запроса
	$result=callInventoryRestMethod 'POST' 'users' "migrate?id=$($source.id)&target=$($destination.id)"
	if ($result -is [bool]) {
		warningLog("migration of employment #$($source.id) -> #$($destination.id) failed")
	} else {
		spooLog("employment data migrated #$($source.id) -> #$($destination.id)")
	}
}

#внутренний номер телефона, привязанный к конкретной кадровой записи ('' если номера нет)
function FetchEmploymentPhone() {
	param
	(
		$id
	)
	$phone=callInventoryRestMethod 'GET' 'phones' 'search-by-user' @{id=$id} $true
	#404/ошибка запроса приезжают как $false
	if ($phone -is [bool]) {return ''}
	$phone=([string]$phone).trim('"')
	if ($phone -eq 'null') {return ''}
	return $phone
}

#обработка пользователя
function ParseUser() {
	param
	(
		[object]$user
	)
	#Выставляем флажки
	#Обновлять пользоватея в АД не надо
	$needUpdate = $false
	#Переименовывать пользователя в АД не надо
	$needRename = $false
	#Увольнять пользователя в АД не надо
	$needDismiss = $false

	
	$invUser = FindUser($user)
	
	#Если пользователь не нашелся
	if ($invUser -eq "error") {
		debugLog($user.sAMAccountname+": Skip: got SAP error")
		return
	}

	#учетку перепривязали на другое трудоустройство - утаскиваем туда же все нажитое
	#(делаем это до чтения телефона, чтобы он читался уже с новой записи)
	MigrateEmployment $global:userRebindFrom $invUser
	
	#проверка увольнения
	#увольняем только если ни одно трудоустройство человека уже не действует
	#(приоритетное выбрано в FindUser, действующее всегда бьет уволенное)
	if ($invUser.Uvolen -eq "1") {
		#смотрим когда уволен
		$resign_date=parseInventoryDate $invUser.resign_date
		if ($null -eq $resign_date) {
			#без внятной даты увольнения учетку не трогаем
			warningLog($user.sAMAccountname+": dismissed in inventory (#"+$invUser.id+"), but resign date ["+$invUser.resign_date+"] is empty or not in ["+$inventory_dateformat+"] format - not deactivating")
		} elseif ((Get-Date) -gt $resign_date) {
			#уже уволен?
			$needDismiss = $true
		}
	}
	
	#
	if ($needDismiss) {
		#Уволенных увольняем
		#проверяем исключения уволенных
		if ($auto_dismiss_exclude -eq $user.sAMAccountname) {
			spooLog($user.sAMAccountname+ ": user dissmissed! Deactivation disabled (exclusion list)!")
		} else {
			if ($auto_dismiss) {
				spooLog($user.sAMAccountname+ ": user dissmissed! Deactivating")
				if ($dismiss_script) {
					Start-Process -FilePath $dismiss_script -ArgumentList $user.sAMAccountname -NoNewWindow
				} else {
					DisableADUser($user)
				}
				return
			} else {
				spooLog($user.sAMAccountname+ ": user dissmissed! Deactivation needed!")
			}
		}
		
	} 
	
	
	#проверка пользователя на совпадение "названия" с ФИО
	if (
		($user.name -ne $invUser.Ename) -or
		($user.cn -ne $invUser.Ename)
	){
		spooLog($user.sAMAccountname+": got AD Name ["+$user.displayName+"] instead of ["+$invUser.Ename+"] - Object rename needed")
		$needRename = $true
	}


	#проверка Выводимого имени пользователя на совпадение с ФИО
	if (
		($user.displayName -ne $invUser.Ename) 
	){
		spooLog($user.sAMAccountname+": got AD displayName ["+$user.displayName+"] instead of ["+$invUser.Ename+"]")
		$user.displayName=$invUser.Ename
		$needUpdate = $true
	}

	#Грузим Ф И О по оттдельности
	$fn=($invUser.fn).trim()
	$mn=($invUser.mn).trim()
	$ln=($invUser.ln).trim()
	$gn=($fn+" "+$mn).trim()

	#проверка Имени и Фамилии пользователя на совпадение с Именем и Фамилией
	if (
		($user.givenName -ne $gn) -or
		($user.sn -ne $ln)
	){
		spooLog($user.sAMAccountname+": got AD firstName+lastName ["+($user.givenName+" "+$user.sn).Trim()+"] instead of ["+($gn+" "+$ln).Trim()+"]")
        if ($ln.Length -gt 0) {
	    	$user.sn=$ln
    		$needUpdate = $true
        } else {
        	if ($write_AD) {
            	$tmpUser = Get-ADUser $user.DistinguishedName
				Set-AdUser $tmpUser -Clear sn
			}
        }
        if ($gn.Length -gt 0) {
    		$user.givenName=$gn
        } else {
        	if ($write_AD) {
            	$tmpUser = Get-ADUser $user.DistinguishedName
				Set-AdUser $tmpUser -Clear givenName
			}
        }
	}

	#Подразделение
	$department=$invUser.orgStruct.name
	if ($department.Length -gt 64) {
		#ограничение длины поля
		$department=$department.Substring(0,64)
	}
	if (
		($department.Length -gt 0) -and
		($user.department -ne $department)
	){
		spooLog($user.sAMAccountname+": got AD Department ["+$user.department+"] instead of ["+$department+"]")
		$user.department=$department
		$needUpdate = $true
	}

	#Организация
	if (
		($invUser.org.uname.Length -gt 0 ) -and
		($user.company -ne $invUser.org.uname )
	){
		spooLog($user.sAMAccountname+": got AD Org ["+$user.company+"] instead of ["+$invUser.org.uname+"]")
		$user.company=$invUser.org.uname
		$needUpdate = $true
	}

	#Должность
	$title=$invUser.Doljnost
	if ($title.Length -gt 128) {
		#ограничение длины поля
		$title=$title.Substring(0,128)
	}
	if (
		($title.Length -gt 0) -and
		($user.title -ne $title)
	){
		spooLog($user.sAMAccountname+": got AD Title ["+$user.title+"] instead of ["+$title+"]")
		$user.title=$title
		$needUpdate = $true
	}

	#uid
	if (($user.adminDescription -ne $invUser.uid) -and ($invUser.uid.Length -gt 0)){
		spooLog($user.sAMAccountname+": got AD UID ["+$user.adminDescription+"] instead of ["+$invUser.uid+"]")
		$user.adminDescription=$invUser.uid
		$needUpdate = $true
	}
		
	#ID организации
	if ($multiorg_support -and ($user.EmployeeNumber -ne $invUser.org_id) -and ($invUser.org_id.Length -gt 0)){
		spooLog($user.sAMAccountname+": got AD Org ID ["+$user.EmployeeNumber+"] instead of ["+$invUser.org_id+"]")
		$user.EmployeeNumber=$invUser.org_id
		$needUpdate = $true
	}
	
	#табельный номер
	if (($user.EmployeeID -ne $invUser.employee_id) -and ($invUser.employee_id.Length -gt 0)){
		spooLog($user.sAMAccountname+": got AD Numbr ["+$user.EmployeeID+"] instead of ["+$invUser.employee_id+"]")
		$user.EmployeeID=$invUser.employee_id
		$needUpdate = $true
	}
		
	#мобильный номер телефона
	$correctedMobile= correctPhonesList($invUser.Mobile)
	if ([string]$user.mobile -ne [string]$correctedMobile) {
		#для поля мобильного делаем обработку на случай если оно стало пустым, т.к. это реальная ситуация, а запись в АД пустого значения делается через задницу
		if ($correctedMobile -eq "") {
			spooLog($user.sAMAccountname+": got AD mobile ["+$user.mobile+"] instead of [empty]")
			if ($write_AD) {
				$tmpUser = Get-ADUser $user.DistinguishedName
				Set-AdUser $tmpUser -Clear mobile
			}
		} else {
			spooLog($user.sAMAccountname+": got AD mobile ["+$user.mobile+"] instead of ["+$correctedMobile+"]")
			$user.mobile=$correctedMobile
			$needUpdate = $true
		}
	}

	#городской номер телефона
	$correctedPhone=correctMobile($user.telephoneNumber)
	if ([string]$user.telephoneNumber -ne [string]$correctedPhone) {
		spooLog($user.sAMAccountname+": got AD telephoneNumber format ["+$user.telephoneNumber+"] instead of ["+$correctedPhone+"]")
		$user.telephoneNumber=$correctedPhone
		$needUpdate = $true
	}

	#Внутренний номер телефона
	#Запрашиваем номер телефона, привязанный к пользователю в Инвентаризации
	$invUserPh=FetchEmploymentPhone $invUser.id
	#если нужно почистить телефон	
	if (($invUserPh -eq "") -and ($user.Pager.Length -gt 0)) {
		spooLog($user.sAMAccountname+": got AD Phone ["+$user.pager+"] instead of ["+$invUserPh+"]")
		if ($write_AD) {
			$tmpUser = Get-ADUser $user.DistinguishedName
			Set-AdUser $tmpUser -Clear Pager
		}					
	} elseIf (
		($invUserPh.length -gt 2 ) -and 
		([string]$invUserPh -ne [string]$user.Pager)
	) {
		spooLog($user.sAMAccountname+": got AD Phone ["+$user.pager+"] instead of ["+$invUserPh+"]")
		$user.pager=$invUserPh
		$needUpdate = $true
	}

	#Почта
	if ([string]$user.mail -ne [string]$invUser.Email) {
		$ad_mail=$false		#проверка что почта в Exchange
		foreach ($dom in $exchange_domains) {
			if (([string]$user.mail).ToLower().EndsWith("@$($dom)".Tolower())) {
				$ad_mail=$true;
			}
		}

		if ($ad_mail) {
			spooLog($user.sAMAccountname+": got Inventory email ["+$invUser.Email+"] instead of ["+$user.mail+"]")
			pushUserData $invUser.id Email $user.mail
		} else {
			spooLog($user.sAMAccountname+": got AD email ["+$user.mail+"] instead of ["+$invUser.Email+"]")
			$user.mail=$invUser.Email
			$needUpdate = $true
		}
	}


	if ($needUpdate) {
		if ($write_AD) {
            #убираем пустые поля
            #$user.PSObject.Properties | ForEach-Object {
                #$_.Name+":"+$_.Value
                #if ($_.Value -eq $null) {
                    #$user.PSObject.Properties.Remove($_.Name)
                    #$user=($user | Select-Object -Property * -ExcludeProperty $_.Name)
                    #spooLog ("Removing $($_.Name) property");
                #}
            #}
			$user 
			Set-AdUser -Instance $user 
			spooLog($user.sAMAccountname+": changes pushed to AD")
		} else {
			spooLog($user.sAMAccountname+": AD push skipped: AD RO mode")
		}
		#exit(0)
	}
	if ($needRename) {
		if ($write_AD) {
			spooLog($user.sAMAccountname+": AdObject renaming to "+$invUser.Ename)
			Rename-AdObject -Identity $user -Newname $invUser.Ename
			spooLog($user.sAMAccountname+": AdObject renamed to "+$invUser.Ename)
		} else {
			spooLog($user.sAMAccountname+": rename $($user.sAMAccountname) -> $($invUser.Ename) skipped: AD RO mode")
		}
	}

	#push данных обратно в БД
	if ([string]$user.sAMAccountname -ne [string]$invUser.Login) {
		spooLog($user.sAMAccountname+": got SAP Login ["+$invUser.Login+"] instead of ["+$user.sAMAccountname+"]")
		pushUserData $invUser.id Login $user.sAMAccountname
	}
	
	
	
}

Import-Module ActiveDirectory

#группы аффилированных организаций: не заданы - каждая организация сама по себе,
#учетка ходит только между трудоустройствами внутри своей организации
$affiliatedOrgs=NormalizeOrgGroups $affiliatedOrgs

if ($args.Length -gt 0) {
	$users = Get-ADUser $args[0] -properties Name,cn,sn,givenName,DisplayName,sAMAccountname,company,department,title,employeeNumber,employeeID,mail,pager,mobile,telephoneNumber,adminDescription

	foreach($user in $users) {
		#подбираем параметры того OU, в котором лежит учетка (нужны для увольнения и приоритетной организации)
		$stickyOrg=$stickyOrgDefault
		foreach ($params in $inventory2ad_sync) {
			if ($user.DistinguishedName -like "*$($params.u_OUDN)") {
				$u_OUDN=$params.u_OUDN
				$f_OUDN=$params.f_OUDN
				if ($null -ne $params.stickyOrg) {$stickyOrg=$params.stickyOrg}
			}
		}
		ParseUser ($user)
	}
} else {
	foreach ($params in $inventory2ad_sync) {
		$u_OUDN=$params.u_OUDN
		$f_OUDN=$params.f_OUDN
		#приоритетная организация может задаваться на каждый OU отдельно
		$stickyOrg=$stickyOrgDefault
		if ($null -ne $params.stickyOrg) {$stickyOrg=$params.stickyOrg}
		$users = Get-ADUser -Filter {enabled -eq $true} -SearchBase $u_OUDN -properties Name,cn,sn,givenName,DisplayName,sAMAccountname,company,department,title,employeeNumber,employeeID,mail,pager,mobile,telephoneNumber,adminDescription
		$u_count = $users | measure 
		Write-Host "Users to sync: " $u_count.Count

		foreach($user in $users) {
			#$user
			ParseUser ($user)
			#exit
		}
	}
}
