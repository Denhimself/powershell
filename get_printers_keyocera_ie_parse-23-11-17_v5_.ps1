Get-Variable -Exclude PWD,*Preference | Remove-Variable -EA 0

function decodingPass ($jid, $encodePass) {
    
    for($i=0; $i -lt $encodePass.Length; $i += 4) {
        
        $result += [char]([Convert]::ToInt32($encodePass.Substring($i, 4),16) -bxor ([char]($jid[$i/4 % ($jid.Length)])))

    }

    return $result

} 

function busyIE($ie) {

    if ($ie.busy) {

        do { Start-Sleep -Milliseconds 1000} while ($ie.ReadyState -ne 4) 

    }

}

function startIE() { 

    do {

        try {
   
            $ie = New-Object -com internetexplorer.application;
       
        } 
        catch [System.Runtime.InteropServices.COMException] {
            Write-Host "Error creating new com-object" # Иногда при частом создание com-объектов выкидывает ошибку.
        }
    }

    until ($ie.ReadyState -eq 0)
  
    #$ie.visible = $true #Сделать Internet Explorer видимым 
    $ie.Silent = $true 
    BusyIE $ie2
    return $ie

}

function closeIE($ie) {

    [void]$ie.Quit()
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($ie) 
    
}

function switchProxyState {
    # Переключает состояние проксисервера на противоположное. Предполагается, что прокси сервер включен. 

    $registryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\"
    $Name = "ProxyEnable"
    $currentValue = Get-ItemProperty -Path $registryPath -Name $Name

    if ($currentValue.ProxyEnable -eq '1') {

        Set-ItemProperty -Path $registryPath -Name $Name -Value '0' # прокси отключен

    } else {

        Set-ItemProperty -Path $registryPath -Name $Name -Value '1' # прокси включен

    }

}


Function parse_Keyocera_total($keyoceraIP) {

    $url = "http://$keyoceraIP" 
    $username="Admin" 
    $password= decodingPass $username 012C0105011E011D010B01330154015D 

    $ie = startIE

    do { # проверяем загрузилась ли страница с фреймом (присутствует документ с фреймом), если нет - загружаем страницу снова.

        $ie.navigate($url);
    
        busyIE $ie

        $frameDoc = $ie.Document.frames[0].document
 
    }

    until ($frameDoc.IHTMLDocument3_getElementById("arg01_UserName"))

    $frameDoc.IHTMLDocument3_getElementById("arg01_UserName").value = $username
    $frameDoc.IHTMLDocument3_getElementById("arg02_Password").value = $password

    $inputTags = $frameDoc.IHTMLDocument3_getElementsByTagName("input")
    $inputTags[112].Click() # нажать на кнопку, которая выполнена тегом "input" со значением "Войти в систему"
    busyIE $ie
    
    $ie2 = startIE 
    closeIE $ie # закрываем первое окно 

    do { 
            
            $ie2.Navigate("$url/rus/index.htm"); # Загружаем ссылку во второй Com-объекте iexplore, если попытаться загрузить в первом, то теряется возможность взаимодействовать с Com-объектом.
    
            busyIE $ie2

            $frameDoc = $ie2.Document.frames[1].document 
 
        }
    
    until ($frameDoc) # проверяем что загрузилась страница с фреймом, если нет - загружаем страницу снова.

    $frameDoc.IHTMLDocument3_getElementById("parentcounters").Click(); # Нажать на ссылку со значением "Счетчик" 
    busyIE $ie2

    $tagsFont = $frameDoc.IHTMLDocument3_getElementsByTagName("font") # В теги <font> находится кол-во распечатанных страниц
    [string]$result = $tagsFont[0].innerText # Всего распечатанных страниц
    
    closeIE $ie2    
    return $result

}

Function parse_hp_total($hpIP) {
    try
       { 
            $uri = "http://$hpIP/hp/device/InternalPages/Index?id=UsagePage" # запрос страницы расхода "HP LaserJet 600 M603"
            $html = Invoke-WebRequest -Uri $uri
            #return [int]$html.ParsedHtml.getElementById("UsagePage.EquivalentImpressionsTable.Print.Total").innerText
            return [int]$html.AllElements.FindById("UsagePage.EquivalentImpressionsTable.Print.Total").innerText
       }

    Catch [System.Net.WebException]
       {
            $uri = "http://$hpIP/hp/device/this.LCDispatcher?nav=hp.Usage" # запрос страницы расхода "HP LaserJet P4015"
            $html = Invoke-WebRequest -Uri $uri
            return [int]($html.ParsedHtml.getElementById("tbl-9160").outerText).split("`n")[5].Substring(5)
       }
}

function refresh-system {
    # The registry changes aren't seen until system is notified about it.
    # Without this function you need to open Internet Settings window for changes to take effect. See http://goo.gl/OIQ4W4
    # Обновляет систему для того, чтобы применились настройки выставляемые в реестре функцией switchProxyState
         
    $signature = @'
    [DllImport("wininet.dll", SetLastError = true, CharSet=CharSet.Auto)]
    public static extern bool InternetSetOption(IntPtr hInternet, int dwOption, IntPtr lpBuffer, int dwBufferLength);
'@

    #$INTERNET_OPTION_SETTINGS_CHANGED   = 39
    $INTERNET_OPTION_REFRESH            = 37
    $type = Add-Type -MemberDefinition $signature -Name wininet -Namespace pinvoke -PassThru
    #$a = $type::InternetSetOption(0, $INTERNET_OPTION_SETTINGS_CHANGED, 0, 0)
    $b = $type::InternetSetOption(0, $INTERNET_OPTION_REFRESH, 0, 0)
    #return $a -and $b
    return $b

}
    
# --------------      Main      ---------------------

$ipAndModel = [Ordered]@{"10.0.0.36"="ip36_HP603";
                         "10.0.0.10"="ip10_HP603";
                         "10.0.0.11"="ip11_HP603";
                         "10.0.0.16"="ip16_HP603";
                         "10.0.0.69"="ip69_HP4015";
			 "10.0.0.82"="ip82_HP603";
		         "10.0.0.33"="ip33_FS4300";
                         "10.0.0.39"="ip39_FS4300";
			 "10.0.0.68"="ip68_FS4300";}

$allPrintersTotal = $null 

foreach ($ip in $ipAndModel.Keys.GetEnumerator()) {[Array]$ipAddresses += $ip} # собираем все IP-адреса 

switchProxyState # Предполагается, что прокси сервер включен.

foreach ($IP in $ipAddresses) {
    # проверяем доступность принтеров по сети

    $stopWarning = $false

    if (!(Test-Connection $IP -Count 1 -Quiet)) {
        
        $stopWarningFlag = $true
        [array]$unreachebaleHosts += $IP
        "{0} : {1}" -f $IP, "Нет пинга!`n"
        
    } 

}

if ($stopWarningFlag) { 
    # Выводим сообщение, если хоть один принтер недоступен, т.е. $stopWarningFlag = $true
     
    $answer = Read-Host 'Наберите "y", чтобы продолжить сбор данных или введите любой символ для выхода'

    if (!($answer -eq "y")) {
    
        exit

    }
}


foreach ($IP in $ipAddresses) {
    # Перебираем список айпишников, пропускаем недоступные, по модели выбираем функцию для парсинга
    
    if ($unreachebaleHosts -contains $IP) {
    
        continue
        
    }
    
    $model = $ipAndModel.get_item("$IP")
    
    if ($model.split("_")[1] -eq "HP603") {
        $total = (parse_hp_total $IP)
    }

    ElseIf ($model.split("_")[1] -eq "HP4015") {
        $total = (parse_hp_total $IP)
    } 

    ElseIf ($model.split("_")[1] -eq "FS4300") { 
         
        [string]$total = (parse_keyocera_total $IP)

    }

    $currentPrinterTotal = "{0}`t {1}: `t{2}`n" -f $model,$IP,$total # Выводим данные в консоль
    #$currentPrinterTotal
    $allPrintersTotal += $currentPrinterTotal

}

$currentDirectory = $PSScriptRoot
$currenDateTime = Get-Date -UFormat "%d-%m-%Y_%H-%M"
Set-Content "$currentDirectory\$currenDateTime.txt" $allPrintersTotal 

#Write-Host 'Variable $allPrintersTotal: '`n
$allPrintersTotal

if (Get-Process iexplore -EA 0) {kill -Name iexplore}

switchProxyState
refresh-system

Write-Host -NoNewLine 'Press any key to continue...';
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown'); 

exit
