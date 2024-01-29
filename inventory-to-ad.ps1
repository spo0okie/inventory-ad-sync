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

. "$($PSScriptRoot)\..\config.priv.ps1"
. "$($PSScriptRoot)\..\libs.ps1\lib_funcs.ps1"
. "$($PSScriptRoot)\..\libs.ps1\lib_inventory.ps1"
. "$($PSScriptRoot)\..\libs.ps1\lib_usr_ad.ps1"


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

#загрузить пользователя из Инвентаризации через REST API
function FindUser() {
	param
	(
		[object]$user
	)

	#Если у нас есть только табельный - считаем что организация=1
    $org_id=$user.employeeNumber
    if (($org_id.Length -eq 0) -or ( -not $mutiorg_support)) {
        #если организация не заявлена, то первая
        #эта ситуация скорее всего возникнет при переходе от инвентаризации версии под одну организацию
        #к инвентаризации версии под множество. Когда БД уже с учетом организаций а АД еще нет
        $org_id=1
    }

	$expand='ln,mn,fn,orgStruct,org'
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
	if (($invUser -isnot [PSCustomObject]) -and ($user.displayName.Length -gt 0)) {$invUser=getInventoryObj 'users' '' @{
		name=$user.displayName;
		expand=$expand;
	}}
	
	#Последний сценарий - табельник вместо ФИО - тут по идее должен находиться один, т.к. такой механизм нужно применять только при сквозной нумерации табельников
	if (($invUser -isnot [PSCustomObject]) -and ($user.displayName.Length -gt 0)) {$invUser=getInventoryObj 'users' '' @{
		num=$user.displayName;
		expand=$expand;
	}}
	
	#если пользователь уволен, и мы искали его по табельнику и у нас есть его uid, то надо поискать по UID (вдруг другая должность есть)
	if (($invUser -is [PSCustomObject]) -and ($invUser.Uvolen -eq "1")) {
		$uid='';
		
		#если UID проставлен в АД - берем оттуда, иначе из инвентори
		if ($user.adminDescription) {
			$uid=$user.adminDescription
		} elseif ($invUser.uid) {
			$uid=$invUser.uid
		}
		
		spooLog "Searching $($user.displayName) is dismissed, searching other employments ($uid)..."
		if ($uid) {$invUser=getInventoryObj 'users' '' @{
			uid=$uid;
			expand=$expand;
		}}
	}

	#Если все-таки не нашли
	if ($invUser -isnot [PSCustomObject]) {
        warningLog("user ["+$user.sAMAccountname+"] with Name ["+$user.displayName+"] - not found in inventory")
        return 'error'
    }

	return $invUser
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
	
	#проверка увольнения
	if ($invUser.Uvolen -eq "1") {
		#смотрим когда уволен
		if ($invUser.resign_date.Length -gt 0) {
			$resign_date=[datetime]::parseexact($invUser.resign_date, $inventory_dateformat, $null)
			#уже уволен?
			if ((Get-Date) -gt $resign_date) {
				$needDismiss = $true
			}
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
				DisableADUser($user)
				#Start-Process -FilePath $dismiss_script -ArgumentList $user.sAMAccountname -NoNewWindow
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
	if (
		($invUser.orgStruct.name.Length -gt 0) -and
		($user.department -ne $invUser.orgStruct.name )
	){
		spooLog($user.sAMAccountname+": got AD Department ["+$user.department+"] instead of ["+$invUser.orgStruct.name+"]")
		$user.department=$invUser.orgStruct.name
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
		($title -gt 0) -and
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
	$invUserPh=callInventoryRestMethod 'GET' 'phones' 'search-by-user' @{id=$invUser.id} $true
	$invUserPh=$invUserPh.trim('"')
    if ($invUserPh -eq 'null') {$invUserPh=''}
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

if ($args.Length -gt 0) {
	$users = Get-ADUser $args[0] -properties Name,cn,sn,givenName,DisplayName,sAMAccountname,company,department,title,employeeNumber,employeeID,mail,pager,mobile,telephoneNumber,adminDescription
} else {
	foreach ($params in $inventory2ad_sync) {
		$u_OUDN=$params.u_OUDN
		$f_OUDN=$params.f_OUDN
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

