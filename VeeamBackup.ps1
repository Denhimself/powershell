Param(
    [string]$serverName,
    [string]$backupStorageFolder,
    [string]$remainTime
)


function get_disk_info ($driveLetter) {
    # Получаем информацию об объёме и свободном месте на диске в гигабайтах

    $disk = get-WmiObject win32_logicaldisk | ? {$_.DeviceID -eq "$($driveLetter):"}
    $disk.Size = [math]::round($disk.Size / 1GB)
    $disk.FreeSpace = [math]::round($disk.FreeSpace / 1GB)

    return $disk

}

function build_spacebar($diskInfo) {
    # Строим полосу занимаемого пространства
    
    $spaceBarSymbol = "▓"
    $fullSize = $diskInfo.Size
    $freeSpace = $diskInfo.FreeSpace
    $occupideSpace = $fullSize - $freeSpace
    $percentPerSpaceBarSymbol = 5
    $totalSteps = 20 # Один шаг это один символ ░
    $stepsCount = [math]::round(($occupideSpace * 100 / $fullSize / $percentPerSpaceBarSymbol))


    $stepsForEmpty = $totalSteps - $stepsCount 

    for ($i = 1; $i -le $stepsForEmpty; $i++) {
        # Кол-во символов "#" обозначающих пустое место.

        $diskEmptySpaceTotalBar += $spaceBarSymbol

    }

    for ($i = 1; $i -le $stepsCount; $i++) {
        # Кол-во символов "#" обозначающих занятое место.

        $diskOccupiedSpaceTotalBar += $spaceBarSymbol

    }

    $spaceBar = ($diskOccupiedSpaceTotalBar,$diskEmptySpaceTotalBar)
    return $spaceBar

}

function set_spacebar_colors ($SpaceBarSymbols, $freeSpace) {
    # Добавлем цвет для полосы занимаемого пространства
     
    if ($freeSpace -le 90) {

        #(Write-host -NoNewline -ForegroundColor Red $SpaceBarSymbols[0]) + (Write-host -ForegroundColor Gray $SpaceBarSymbols[1])
        '<font color="red" size="4">' + $SpaceBarSymbols[0] + '</font>' + '<font color="8a7f8e" size="4">' + $SpaceBarSymbols[1] + '</font>'

    } elseif ($freeSpace -le 300) {

        #(Write-host -NoNewline -ForegroundColor Yellow $SpaceBarSymbols[0]) + (Write-host -ForegroundColor Gray $SpaceBarSymbols[1]) 
        '<font color="yellow" size="4">' + $SpaceBarSymbols[0] + '</font>' + '<font color="8a7f8e" size="4">' + $SpaceBarSymbols[1] + '</font>'

    } else {

        #(Write-host -NoNewline -ForegroundColor Green $SpaceBarSymbols[0]) + (Write-host -ForegroundColor Gray $SpaceBarSymbols[1])
        '<font color="green" size="4">' + $SpaceBarSymbols[0] + '</font>' + '<font color="8a7f8e" size="4">' + $SpaceBarSymbols[1] + '</font>'
    }
}


function sendMail ($subject,$body){

     #SMTP server name
     $smtpServer = "mail.company.net"

     #Creating a Mail object
     $msg = new-object Net.Mail.MailMessage

     #Creating SMTP server object
     $smtp = new-object Net.Mail.SmtpClient($smtpServer)

     $msg.isBodyHTML = $true

     #Email structure 
     $msg.From = “asaf@company.net“
     $msg.To.Add(“adfaf@company.net“)

     $msg.subject = $subject
     $msg.body = $body 

     #Sending email 
     $smtp.Send($msg)
  
}

$jobSpecList = @("JobId", "JobName", "VmName", "ObjectId", "FolderName", "CompressionLevel", "VmToolsQuiesce", "RetainDatetime","ModifiedBy") 
$auxDataList = @("BackupSize", "DataSize", "CompressRatio", "WorkDuration")


function get_backup_info ($backupInfo, $paramList, $xPath) {


    foreach ($param in $paramList) {
         
        $paramValue =  (Select-Xml -Content $backupInfo.$xPath -XPath "//$param").Node.InnerText

        if ($param -eq "VmName") {

            $paramValue = '<font color="6600ff">' + ">>>$paramValue<<<" + '</font>'

        }
        
        if ($param -eq "BackupSize" -or $param -eq "DataSize") {

            $paramValue = '<font color="green">' + [string]([math]::round($paramValue/[math]::Pow(1024, 3),3))  + " GB</font>"
        }
        
        $param + " : " + $paramValue + "<br>"
    }

}



function set_result_color ($resultValue ) {

    if ($resultValue -eq "success") {

        "Result : " + '<font color="green" size="6">' + $resultValue + "</font><br>"

    } elseif ($resultValue -eq "warning") {

        "Result : " + '<font color="f3da0b" size="6">' + $resultValue + "</font><br>"

    } else {
    
        "Result : " + '<font color="red" size="6">' + $resultValue + "</font><br>"

    }
}



function check_backup_files_limit ($backupFolder, $limit) {
    
    $countBackUps = (Get-ChildItem $backupFolder -Name "*.vbk" | Measure-Object ).Count;

    if ($countBackUps -gt $limit) {
        
         $result =  '<font color="red" size="4">' + "WARNING! Backup folder contains $countBackUps backups! Files are not delete!" + '</font>'

    } else {

         $result = '<font color="green" size="4">' + "Backup folder contains $countBackUps backups." + '</font>'

    }

    #$return
    return $result 
}


function delete_backups_over_limit ([int]$backupLimit, $folderPath) {
    # удаляем все файлы в папке с бекапами определенного сервера, сверх кол-ва указаного в переменной $backupLimit

    [int]$countBackUps = (Get-ChildItem $folderPath -Name "*.vbk" | Measure-Object ).Count;

    if ($countBackUps - $backupLimit -gt 0) {

        $filesList = Get-ChildItem $folderPath -Name "*.vbk" | select -First ($countBackUps - $backupLimit)

        foreach ($file in $filesList) {

                Remove-Item "$folderPath\$file"
        }
    }
 }


###############   Создаем бэкап  #############################################################################

$vm = Find-VBRHvEntity -Name $serverName
#$backupInfo = Start-VBRZip  -Folder $backupStorageFolder -Entity $vm -Compression 5 -AutoDelete $remainTime
$backupInfo = Start-VBRZip  -Folder $backupStorageFolder -Entity $vm -Compression 5 
###############  Определяем место на диске и делаем цветную полосу используемого пространства  ###############

$diskInfo = get_disk_info "D"
$totalSpaceBarSymbols = build_spacebar $diskInfo
$coloredSpaceBar = set_spacebar_colors $totalSpaceBarSymbols $($diskInfo.FreeSpace) 

###############  Выставляем предел кол-ва бэкапов в папке и получаем список всех бэкапов в данной папке  #####

$backupFilesLimit = 3
delete_backups_over_limit $backupFilesLimit $backupStorageFolder
$folderFilesList = (Get-ChildItem $backupStorageFolder -Filter *.vbk)

###############  Создаем письмо с отчетом и отправляем  ######################################################

$subject = "Veeam_Backup-$($backupInfo.JobName) $($backupInfo.Result) $($backupInfo.BaseProgress)"
$body = (set_result_color $backupInfo.Result) +`
        (get_backup_info $backupInfo $jobSpecList "JobSpec") + `
        (get_backup_info $backupInfo $auxDataList "AuxData") + ` 
        "CreationTime" + " : " + '<font color="green">' + [string]($backupInfo.CreationTime) + "</font><br>" +`
        "EndTime" + " : " + '<font color="green">' + [string]($backupInfo.EndTime) + "</font><br>" +
        "<br>Server free space   " + $coloredSpaceBar + "     <b>$($diskInfo.FreeSpace)</b>Gb/<b>$($diskInfo.Size)</b>Gb<br>" +
        "<br>" + (check_backup_files_limit $backupStorageFolder $backupFilesLimit) + "<br>" +
        ($folderFilesList | % {'<br><font face="Helvetica" size="2">{0}&emsp;{1}&emsp;{2:N2}' -f $_.Name, $_.CreationTime, ($_.Length / 1mb) + " MB</font>"}) 

sendMail $subject $body



