
#Функция получения авторизационного ключа сессии, out = authkey
##########################################################
Function Get-ZAPIauthKey ($baseZURL,$Credentials)
{
    $params = @{
        body = @{
            "jsonrpc"= "2.0"
            "method"= "user.login"
            "params"= @{
                "user"= $Credentials.UserName -replace "DPC\\","" -replace "dpc\\"
                "password"= $Credentials.GetNetworkCredential().Password
            }
            "id"=1
            "auth"= $null
        }|ConvertTo-Json
        uri="$baseZURL/api_jsonrpc.php"
        headers = @{"Content-Type"="application/json"}
        method = "Post"
    }
    try {
        $json = Invoke-WebRequest @params
        return ($json|ConvertFrom-Json).result
    } catch {
        write-host "Error!!!" -f Red
        #Добавить вывод конкретной ошибки
        write-host $_ -f Red
    }
    
}
##########################################################

#Функция получения списка групп хостов, out(groupid,name)
##########################################################
Function Get-ZAllHostGroups ($ZAuthKey, $baseZURL)
{
    $params = @{
        body = @{
            "jsonrpc"= "2.0"
            "method"= "hostgroup.get"
            "params"= @{
                output= ("groupids","name")
            }
            "id"=2
            "auth"= $ZAuthKey
        }|ConvertTo-Json
        uri="$baseZURL/api_jsonrpc.php"
        headers = @{"Content-Type"="application/json"}
        method = "POST"
    }
    try {
        $json = Invoke-WebRequest @params|ConvertFrom-Json
        return $json.result
    } catch {
        write-host "Error!!!" -f Red
        #Добавить вывод конкретной ошибки
        write-host $_ -f Red
    }
}
##########################################################

#Функция получения списка триггеров по группе хостов, out(hostid, host, triggerid, description, expression)
##########################################################
Function Get-ZTriggers($baseZURL, $ZAuthKey, $HostGroup)
{
    $params = @{
        body = @{
            "jsonrpc"= "2.0"
            "method"= "trigger.get"
            "params"= @{
                output= "extend"
                groupids = $HostGroup.groupid
                selectHosts = ("hostid","name")
            }
            "id"=2
            "auth"= $ZAuthKey
        }|ConvertTo-Json
        uri="$baseZURL/api_jsonrpc.php"
        headers = @{"Content-Type"="application/json"}
        method = "POST"
    }
    try {
        $json = Invoke-WebRequest @params|ConvertFrom-Json
        $res = foreach($trigger in $json.result){
            $r = ""|select host,triggerid, description, expression
            $r.host = $trigger.hosts.name
            $r.triggerid = $trigger.triggerid
            $r.description = $trigger.description
            $r.expression = $trigger.expression
            $r
        }
        return $res
    } catch {
        write-host "Error!!!" -f Red
        #Добавить вывод конкретной ошибки
        write-host $_ -f Red
    }
}
##########################################################

#Функция получения списка серверов из AF  out(name,ip,segment,os,ram,proces,hddos,hdddata,cod,active)
Function Get-ServFromAF($Contour, $System) #$lsn, $database) #,$Contour, $Pods)
{
    #Сервер сервисной БД, Имя сервисной БД, Контур
    SWITCH ($Contour )
    {
        "КПЭ" {
            $dbsrv = ''
            $database = ''
        }
        "ППК" {
            $dbsrv = ''
            $database = ''
        }
        "КОЭ" {
            $dbsrv = ''
            $database = ''
        }
        Default {break}
    }

    #Корректировка подсистемы
    SWITCH ($System)
    {
        "" {$System = ""}
        "" {$System = ""}
        "" {$System = ""}
        "" {$System = ""}
        "" {$System = ""}
        "" {$System = ""}
    }
    $Pods = $System 
        
    #Строка подключения к сервисной БД
    $ssrt = "Server=$dbsrv;Database=$database;Integrated Security=True;" 
    
    #Запрос на получение списка серверов
    $select = "SELECT * FROM [$database].[dbo].[servers_cod] WHERE kontur = '$($Contour)' AND subsystem = '$($Pods)' AND active='Активный'" 

    #Получение списка серверов
    $table = Get-MSSQLData -SQLQuery $select -ConnectionString $ssrt
    return $table
}
####################################################################

#Функция получения списка служб сервера из AF  #out(name,ip,segment,os,ram,proces,hddos,hdddata,cod,active)
function Get-SvcFromAF ($Server, $Contour, $System)
{
    #Сервер сервисной БД, Имя сервисной БД, Контур
    SWITCH ($Contour )
    {
        "КПЭ" {
            $dbsrv = ''
            $database = ''
        }
        "ППК" {
            $dbsrv = ''
            $database = ''
        }
        "КОЭ" {
            $dbsrv = ''
            $database = ''
        }
        Default {break}
    }

    #Корректировка подсистемы
    #Корректировка подсистемы
    SWITCH ($System)
    {
        "" {$System = ""}
        "" {$System = ""}
        "" {$System = ""}
        "" {$System = ""}
        "" {$System = ""}
        "" {$System = ""}
    }
    $Pods = $System 
        
    #Строка подключения к сервисной БД
    $ssrt = "Server=$dbsrv;Database=$database;Integrated Security=True;" 
    
    #Запрос на получение списка серверов
    $select = "SELECT [servic_name] FROM [$database].[dbo].[service_win] where machine_name = '$($Server)'"

    #Получение списка серверов
    $table = Get-MSSQLData -SQLQuery $select -ConnectionString $ssrt
    if ($table -ne $null)
    {
        return $table.servic_name
    }
}

####################################################################
#Функция разукрашивания файла Excell
Function Colorite-Excell ($Path)
{
    try
        {
        $excell = New-Object -comobject Excel.Application
        $excell.visible = $False
        $workbook = $excell.workbooks.open($Path)
        $colors = (34,35,36,37,38,39,40)
        $usedcolor = @{"u1"=0;"u2"=1;"u3"=1}
        $usednumber = ''

        foreach ($sheet in $workbook.Sheets)
        {
            Foreach ($row in $sheet.UsedRange.Rows)
            {
                #Окрашивание заголовка
                if ($row.row -eq 1)
                {
                    $row.font.bold=$true
                    $row.interior.colorindex=20
                }
                #Окрашивание строк
                elseif ($sheet.Cells.Item($row.row,1).text -eq $usednumber) 
                {
                    $row.interior.colorindex=$usedcolor.u1
                }
                elseif ($sheet.Cells.Item($row.row,1).text -ne $usednumber)
                {
                    $usedcolor.u3 = $usedcolor.u2
                    $usedcolor.u2 = $usedcolor.u1
                    $usedcolor.u1 = Get-Random -InputObject $($colors|where{($_ -ne $usedcolor.u3) -and ($_ -ne $usedcolor.u2)})
                    $usednumber = $sheet.Cells.Item($row.row,1).text
                    $row.interior.colorindex=$usedcolor.u1
                }
            }
        }
        $workbook.Save()
        $workbook.Close()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)|Out-Null
        Remove-Variable workbook -ErrorAction SilentlyContinue|Out-Null
        $excell.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excell)|Out-Null
        Remove-Variable excell -ErrorAction SilentlyContinue|Out-Null
    }
    catch
    {
        Write-Host $_.Error
    }
}
####################################################################

#Функция получения списка триггеров по выбранной подсистеме из Zabbix # out(host, description, expression)
##################################
Function Get-ZSTriggers($url, $Credentials, $System)
{
    Write-Host "Получаем ключ авторизации Zabbix пользователя" -f Yellow
    #Получаем ключ авторизации
    $AuthKey =Get-ZAPIauthKey -baseZURL $url -Credentials $Credentials
    if ($AuthKey -eq $null)
    {
        Write-Host "Критическая ошибка: не получен ключ авторизации Zabbix" -f Red
        Write-Host "Проверьте правильность сервера Zabbix, пользователя, пароль и попробуйте ещё раз!" -f Yellow
        break
    }
    else
    {
        Write-Host $AuthKey -f Green
    }

    Write-Host "Получаем список серверов подсистемы $System из Zabbix" -f Yellow
    #Получаем имя группы хостов
    #groupid, name
    $HostGroups = Get-ZAllHostGroups -baseZURL $url -ZAuthKey $AuthKey

    #Фильтруем группы по искомой подсистеме
    $HostGroupGP3 = $HostGroups|where {$_.name -like "$Contour*$System"}
    Write-Host $HostGroupGP3 -f Green
    if ($HostGroupGP3 -eq $null) {
        $Triggers = ""
        return $Triggers
    }
    Write-Host "Получаем список триггеров подсистемы $System из Zabbix" -f Yellow
    #Получаем список триггеров
    #host, description, expression
    $Triggers = Get-ZTriggers -baseZURL $url -ZAuthKey $AuthKey -HostGroup $HostGroupGP3
    
    return $Triggers
}
########################################

#Функция запросов в БД MS SQL
###########################################################
Function Global:Get-MSSQLData
    {
        <#
            .SYNOPSIS
                Функция предназначена для получения данных из таблиц БД MS SQL.
            .DESCRIPTION
                Данная функция позволяет выполнять запросы к БД MS SQL типа SELECT.
                Запросы типа UPDATE, DELETE, INSERT, CREATE, DROP данной функцией
                НЕ ПОДДЕРЖИВАЮТСЯ!
                Данные можно передавать в функцию по конвейеру. Также досупен ввод
                данных с клавиатуры.
            .PARAMETER SQLQuery
                SQL-запрос типа SELECT.
                Тип данных - [string[]] (строковый массив).
            .PARAMETER ConnectionString
                Строка подключения к БД MS SQL.
                Тип данных - [string] (строковый).
            .EXAMPLE
                Get-MSSQLData -SQLQuery "Select * From dbo.MyTable" -ConnectionString "Server=MyServer;Database=MyDataBase;Integrated Security=True"

                COLUMN1                        COLUMN2                                                                                                             
                -------                        -------                                                                                                             
                XXXX                           AAAA                                                                                                      
                YYYY                           BBBB                                                                                                               
                ZZZZ                           CCCC 
        #>
        [CmdletBinding `
                        (
                            SupportsPaging = $true,
                            SupportsShouldProcess=$true
                        )]
        Param
            (
                [Parameter `
                            (
                                Mandatory = $true,
                                ValueFromPipeLine=$true,
                                HelpMessage = "MS SQL запрос типа SELECT"
                            )]
                [ValidateScript `
                                (
                                    {
                                        $_ -match "SELECT " `
                                        -and $_ -notmatch "INSERT " `
                                        -and $_ -notmatch "UPDATE " `
                                        -and $_ -notmatch "DELETE " `
                                        -and $_ -notmatch "CREATE " `
                                        -and $_ -notmatch "DROP "
                                    }
                                )]
                [string[]]$SQLQuery,
                [Parameter `
                            (
                                Mandatory = $true,
                                ValueFromPipeLine=$true,
                                HelpMessage = "Строка подключения к БД MS SQL"
                            )]
                [string]$ConnectionString
                
            )
        BEGIN
            {
                #Создается подключение к экземпляру БД
                $conn = New-Object System.Data.SqlClient.SqlConnection($ConnectionString)
            }
        PROCESS
            {
                Write-Verbose "Попытка установить соединение с БД..."
                    #Открытие соединения с БД
                    try
                        {
                            $conn.Open()
                            $status = $true
                        }
                    catch
                        {
                            $err = $Global:Error[0]
                            $status = $false
                            Write-Host "Возникла ошибка при попытке открыть соединение с БД!" -ForegroundColor Red
                            Write-Host $err -ForegroundColor Red
                        }
                Write-Verbose "Попытка установить соединение с БД окончена"
                    if ($status -eq $true)
                        {
                            #Создание SQL-запрса
                            $sql = $conn.CreateCommand()
                            $sql.CommandText = $SQLQuery
                            Write-Verbose "Попытка исполнения SQL-запроса на БД..."
                                #Выполнение SQL-запрса
                                try
                                    {
                                        $res = $sql.ExecuteReader()
                                    }
                                catch
                                    {
                                        $err = $Global:Error[0]
                                        $return = $false
                                        Write-Host "Возникла ошибка при попытке выполнить запрос на БД!" -ForegroundColor Red
                                        Write-Host $err -ForegroundColor Red
                                    }
                            Write-Verbose "SQL-запрос успешно выполнен на БД"
                            #Создание таблицы для выгрузки результата
                            $table = New-Object System.Data.DataTable
                            Write-Verbose "Попытка выгрузки результатов выполнения SQL-запроса в таблицу..."
                                #Выгрузка результатов SQL-запроса в таблицу
                                try
                                    {
                                        $table.Load($res)
                                        $return = $table
                                    }
                                catch
                                    {
                                        $err = $Global:Error[0]
                                        $return = $false
                                        Write-Host "Возникла ошибка при попытке выгрузить результаты запроса в таблицу!" -ForegroundColor Red
                                        Write-Host $err -ForegroundColor Red
                                    }
                            Write-Verbose "Выгрузка результатов выполнения SQL-запроса в таблицу завершена"
                            #Закрытие соединения с БД
                            $conn.Close()
                        }
                    else
                        {
                            $return = $status
                        }
            }
        END
            {
                #Возвращаемое значение
                return $return
            }
    }
###########################################################