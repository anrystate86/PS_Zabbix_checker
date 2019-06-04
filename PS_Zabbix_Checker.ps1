Import-Module "$PSScriptRoot\System\ImportExcel"
#Прикручиваем файл с функциями
. $PSScriptRoot\PSZLib.ps1
# Переменные временно вводимые вручную
$Username = "Login"
$Password = "Password"
#---------------------------------
# Ввод логин\пароля в переменную для дальнейшего использования
$Pass = ConvertTo-SecureString -AsPlainText $Password -Force
$MyCredentials = New-Object System.Management.Automation.PSCredential -ArgumentList $Username, $pass

cls

#Параметры для поиска триггеров мониторинга
$Contour = "" 
$System = "" 

SWITCH ($Contour) {
    "" {$Zurl = "http://zbx/zabbix"}
    "" {$Zurl = "http://zbx/zabbix"}
    "" {$Zurl = ("http://zbx/zabbix","http://zbx/zabbix")}
}

#Путь сохранения файла
$OutFilePath = "$PSScriptRoot\Triggers_$($Contour)_$($System)_$(Get-Date -Format "dd.MM.YY").xlsx"

$Triggers = $Zurl|foreach{
    Get-ZSTriggers -url $_ -credentials $MyCredentials -system $System
}

Write-Host "Получаем список серверов подсистемы $System из AF" -f Yellow
$Hosts = Get-ServFromAF -contour $Contour -system $System

Write-Host "Формируем список сервер-триггер" -f Yellow
#Формируем список сервер-триггер
$HostWTrigs = foreach($thost in $Hosts.name){
    $Services = Get-SvcFromAF -Server $thost -Contour $Contour -system $System
    
    $Trigs = $Triggers|where{$_.host -imatch $thost}

    if ($Trigs -ne $null)
    {
        $rh = foreach($trig in $Trigs){
            $rt=''|select host, description, service, serviceaf, triggerid, expression
            $rt.serviceaf = $Services|where {$trig.description -match $_}
            $rt.host = $thost 
            $rt.triggerid = $trig.triggerid
            $rt.description = (($trig.description).Replace("{HOST.NAME}",$rt.host)).replace('{HOST.HOST}',$rt.host)
            $rt.service = ($rt.description|where {$_ -match "служба"})-replace "^\w{2}\s\w+\-\w+\s\w+\s" -replace " в статусе" -replace " {ITEM.VALUE}$" -replace " статусе"
            if ($rt.service -eq "SQL Server (MSSQLSERVER)") {$rt.service = "MSSQLSERVER"}
            $rt.expression = $trig.expression
            $rt
        }
        
        if (($Services|where {$_ -notin ($rh.service -ne $null|select -Unique)}) -eq $null)
        {
            #Write-Host "На сервере $thost мониторинг служб соответствует AF" -f Green
        }
        else
        {
            Write-Host "На сервере $thost мониторинг служб НЕ соответствует AF" -f Red
            Write-Host "    Службы в AF    : $($Services|sort)" -f Yellow
            Write-Host "    Службы в Zabbix: $($rh.service -ne $null|sort)" -f Yellow
            
            $out = ""|select host,serviceaf,servicez
            $out.host = $thost
            $out.serviceaf = ($Services|sort) -join " "
            $out.servicez = ($rh.service -ne $null|sort) -join " "
            $out|Export-Csv -Path "$PSScriptRoot\BadTriggers_$($Contour)_$($System)_$(Get-Date -Format "dd.MM.YY").csv" -Delimiter ';' -NoTypeInformation -Append -Encoding UTF8
        }
    }
    else
    {
        $rh=''|select host, description, service, serviceaf, triggerid, expression
        $rh.host = $thost
        $rh.serviceaf = $Services
        $rh.triggerid = "00000"
        $rh.description = "Триггеры мониторинга отсутсвуют"
        $rh.expression = ""
        if ($rh.serviceaf -ne $null)
        {
            Write-Host "На сервере $thost мониторинг служб НЕ ведется в Zabbix" -f Red
            Write-Host "    Службы в AF    : $($Services|sort)" -f Yellow
            Write-Host "    Службы в Zabbix: Отсутствуют" -f Yellow
            
            $out = ""|select host,serviceaf,servicez
            $out.host = $thost
            $out.serviceaf = ($Services|sort) -join " "
            $out.servicez = "нет"
            $out|Export-Csv -Path "$PSScriptRoot\BadTriggers_$($Contour)_$($System)_$(Get-Date -Format "dd.MM.YY").csv" -Delimiter ';' -NoTypeInformation -Append -Encoding UTF8
        }
    }
    $rh
}
Write-Host "Формируем список оставшихся триггеров, не связанных с серверами" -f Yellow
#Формируем список оставшихся триггеров, не связанных с серверами
$otherTrigs = $Triggers|where{$_.triggerid -notin $HostWTrigs.triggerid}

Write-Host "Выгружаем результаты в Excell" -f Yellow
#Выгружаем результаты в Excell
$HostWTrigs|Export-Excel -Path $OutFilePath -WorkSheetname "Triggers" -ClearSheet -TitleBold -AutoSize -FreezeTopRow -KillExcel
$otherTrigs|Export-Excel -Path $OutFilePath -WorkSheetname "OtherTriggers" -ClearSheet -TitleBold -AutoSize -FreezeTopRow -KillExcel

Write-Host "Окрашиваем Excell-файл" -f Yellow
#Окрашивание Excell-файла
Colorite-Excell -path $OutFilePath

""|Export-Excel -Path $PSScriptRoot\System\dumb.xlsx -WorkSheetname 1 -KillExcel

Write-Host "Done!!!" -f Green