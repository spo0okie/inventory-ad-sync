#Привязка компьютеров из AD к сервисам в инвентаризации.
#
#Нужен для динамически создаваемых машин (VDI): AD-объект появляется в своем OU,
#а в инвентаризации ОС надо прицепить к соответствующему сервису VDI.
#
#Работает ТОЛЬКО на добавление: чужие связи (проставленные руками или другим OU) не трогаются,
#машины, пропавшие из OU, от сервиса не отцепляются.
#
#Полный прогон по всем OU из конфига:
#	comps-to-services.cmd
#Один компьютер (основной способ отладки, OU игнорируются):
#	powershell.exe -noprofile -executionpolicy bypass -file comps-to-services.ps1 vdi-0123

#Windows PowerShell 5.1 по умолчанию предлагает серверу только SSL3/TLS1.0 -
#без TLS1.2 все REST-запросы к инвентаризации падают с ошибкой создания защищенного канала
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

. "$($PSScriptRoot)\..\config.priv.ps1"
. "$($PSScriptRoot)\..\libs.ps1\lib_funcs.ps1"
. "$($PSScriptRoot)\..\libs.ps1\lib_inventory.ps1"

#404 (нет такой ОС / нет такого сервиса) - штатная ситуация, о ней пишется своя внятная строка
$global:skip404errors=$true

#кэш сервисов: ссылка из конфига (id или имя) -> объект сервиса ($false - не нашли)
$global:servicesCache=@{}


#запись набора сервисов ОС в инвентаризацию
function pushCompServices() {
	param
	(
		[string]$id,
		$services_ids
	)
	#services_ids объявлен в Comps::linksSchema, PUT с ним перезаписывает весь набор связей,
	#поэтому наверх всегда уходит объединенный список, а не одна добавляемая связь
	$params = @{
		services_ids=@($services_ids);
		id=$id;
	}

	if ($write_inventory) {
		$null = setInventoryData 'comps' $params
	} else {
		spooLog("invPush: skip comp #$id services_ids = $(@($services_ids) -join ',') INV: RO mode")
	}
}


#ищет сервис в инвентаризации: по ID если ссылка из конфига число, иначе по имени
#результат кэшируется, возвращает объект сервиса или $false
function FindService() {
	param
	(
		[string]$ref
	)

	if ($ref.Length -eq 0) {return $false}
	if ($global:servicesCache.ContainsKey($ref)) {return $global:servicesCache[$ref]}

	if ($ref -match '^\d+$') {
		$service=callInventoryRestMethod 'GET' 'services' $ref
	} else {
		#paramsString в lib_inventory параметры не кодирует, а в именах сервисов есть пробелы и двоеточия
		$service=getInventoryObj 'services' ([System.Web.HttpUtility]::UrlEncode($ref))
	}

	if (($service -is [bool]) -or ($null -eq $service.id)) {
		errorLog("service [$ref] not found in inventory")
		$service=$false
	} else {
		debugLog("service [$ref] resolved to #$($service.id) [$($service.name)]")
	}

	$global:servicesCache[$ref]=$service
	return $service
}


#ищет ОС в инвентаризации по FQDN (или по короткому имени, если FQDN в AD не заполнен)
#возвращает объект ОС (с подтянутыми сервисами) или $false
function FindComp() {
	param
	(
		[string]$name
	)
	#expand=services нужен, чтобы узнать уже имеющиеся связи: services_ids через expand не отдается
	$comp=callInventoryRestMethod 'GET' 'comps' 'search' @{name=$name; expand='fqdn,services'}
	if (($comp -is [bool]) -or ($null -eq $comp.id)) {return $false}
	return $comp
}


#обработка одного компьютера: дописывает ему недостающие связи с сервисами
function ParseComp() {
	param
	(
		$adComp,
		$serviceRefs
	)

	#в AD у динамически созданной машины FQDN может быть еще не заполнен - тогда ищем по короткому имени
	$name=[string]$adComp.DNSHostName
	if ($name.Length -eq 0) {$name=[string]$adComp.Name}

	$comp=FindComp $name
	if ($comp -is [bool]) {
		warningLog("$($name): not found in inventory")
		return
	}

	if ($comp.archived) {
		warningLog("$($name): inventory comp #$($comp.id) is archived, skipping")
		return
	}

	#текущий набор сервисов ОС
	$currentIds=@(@($comp.services) | ForEach-Object {[string]$_.id})
	$newIds=@()

	foreach ($ref in @($serviceRefs)) {
		$service=FindService ([string]$ref)
		if ($service -is [bool]) {continue}

		$serviceId=[string]$service.id
		if (($currentIds -contains $serviceId) -or ($newIds -contains $serviceId)) {
			debugLog("$($name): already in service #$serviceId [$($service.name)]")
			continue
		}

		spooLog("$($name): adding comp #$($comp.id) to service #$serviceId [$($service.name)]")
		$newIds+=$serviceId
	}

	if ($newIds.Count -eq 0) {return}

	pushCompServices $comp.id ($currentIds+$newIds)
}


Import-Module ActiveDirectory

if ($args.Length -gt 0) {
	$comps = Get-ADComputer $args[0] -properties DNSHostName

	foreach($comp in $comps) {
		#OU в этом режиме не обходим, но сервисы берем из тех элементов конфига, в чей OU объект попал
		$serviceRefs=@()
		foreach ($params in $comps2services_sync) {
			if ($comp.DistinguishedName -like "*$($params.OUDN)") {$serviceRefs+=@($params.service)}
		}
		if ($serviceRefs.Count -eq 0) {
			warningLog("$($comp.Name): no OU in `$comps2services_sync matches $($comp.DistinguishedName)")
		} else {
			ParseComp $comp $serviceRefs
		}
	}
} else {
	foreach ($params in $comps2services_sync) {
		$comps = Get-ADComputer -Filter {enabled -eq $true} -SearchBase $params.OUDN -properties DNSHostName
		Write-Host "$($params.OUDN): comps to sync: $(@($comps).Count)"

		foreach($comp in $comps) {
			ParseComp $comp @($params.service)
		}
	}
}
