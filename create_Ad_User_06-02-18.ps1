Get-Variable -Exclude PWD,*Preference | Remove-Variable -EA 0

function display_list ($List, $columnAmount, $rowsInColumn) {
    # Печатает в несколько колонок полученную хэш-таблицу.
    # Заполнение идет по столбцам, а не как по дефолту - построчно. 

    [System.Collections.ArrayList]$shuffledList = [ordered]@{}

    for($i = 0; $i -lt $rowsInColumn; $i++){

        for($j = 0; $j -le $columnAmount-1; $j++){

            $index = $j * $rowsInColumn + $i

            if ($index -lt ($List.Length)){

                $shuffledList += $List[$index]

            } else {

                $shuffledList += "" 
                break

            }

        }

    }

    $shuffledList | fw {$_} -Column $columnAmount -Force

}

function get_user_info ($login) {
    # Парсим данные пользователя с oit.company.net

    $html = Invoke-WebRequest -Uri http://portal.company.net/reference/index.phtml -Method POST -Body @{email="$login@company.net"} 

    for ($i = 50; $i -le 55; $i++) {
	# Результатом поиска будет одна запись с искомым пользователем. Данные находятся в тегах <td>, с 50 по 55.   

        [Array]$user_data = $user_data + $html.ParsedHtml.getElementsByTagName("td")[$i].innerText

    } return [ordered]@{FIO=$user_data[0]; 
                        department=($user_data[1]).split(",")[0];
                        title=($user_data[1]).split(",")[1].Trim(); 
                        email=$user_data[2]; 
                        phoneNum=$user_data[3];  
                        innerPhoneNum=$user_data[4]}

}

function set_new_AD_user ($userInfo, $OU, $manager) {
     
    $displayName = $($userInfo.FIO) # !!! временно ставить "test_" перед строкой !!!
    $login = $userInfo.email.Split('@')[0]
    $surName = $displayName.Split(' ')[0]

    $managerAttributes = get_manager $OU

     
    if (is_ou_contain_manager $OU) {
    # Проверяем наличие руководителя в OU, если его нет, то AD-атрибут "Manager" создаваемого пользователя оставим пустым.

        if (is_manager $userInfo.title) {
        # Если создаваемый сотрудник - начальник, то в его AD-атрибут "Manager" присваиваем AD-атрибут "Manager" текущего начальника в данной OU.
        # Обычному сотруднику в AD-атрибут "Manager" присваиваем AD-атрибут "DistinguishedName" текущего начальника из данной OU. 
        
            [string]$manager = $($managerAttributes.manager)

        } else {

            [string]$manager = $($managerAttributes.DistinguishedName)

        }

    } else {

       $manager = $null

    }

    # Создаем пользователя и задаем все нужные данные.
        New-ADUser  -SamAccountName "$login"`
                    -Manager $manager `
                    -AccountPassword (ConvertTo-SecureString "P@ssw0rd" -AsPlainText -Force) `
                    -Path "$OU" `
                    -Name $displayName.Trim()`
                    -City "$($managerAttributes.City)"`
                    -StreetAddress "$($managerAttributes.StreetAddress)"`
                    -Office "$($managerAttributes.Office)"`
                    -OtherAttributes @{Displayname = "$displayName";
                                       GivenName ="$displayName";
                                       sn="$surName";
                                       UserPrincipalName="$login";
                                       telephoneNumber="$($userInfo.phoneNum)";
                                       mail=$userInfo.email;
                                       Company='ООО "Мелкомягкий"';
                                       Title=$userInfo.title;
                                       Department="$($userInfo.department)"} -Enabled $true

}

function show_created_user_info ($verifiableLogin) {
    # Выводит данные созданного пользователя.

    Get-ADUser $verifiableLogin -Properties * | select DistinguishedName, Name, SamAccountName, City, StreetAddress, Office, Surname, telephoneNumber,mail,Company,Title,Department,Manager

}

function get_manager ($OU) {
    # Узнать кто руководитель в данной OU
    
    $manger = Get-ADUser -SearchBase "$OU" -SearchScope OneLevel -Filter {directReports -like "*" -and Enabled -eq "True"} -Properties City, StreetAddress, Office, manager
    return $manger

}

function change_manager ($oldManager, $newManager, $OU) {
    # Меняем руководителя, имя которого получили в $oldManager, каждому юзеру в OU c таким руководителем на нового.

    get-aduser -f {manager -eq $oldManager} -SearchBase $OU -SearchScope OneLevel | Set-ADUser -Manager "$newManager"

}

function is_manager ($userTitle) {
    # По должности проверяем является ли создаваемый юзер начальником или руководителем

    if ($userTitle -like "*начальник*" -or ($userTitle -like "*руководитель*")) {

        return $true

    } else {

        return $false

    }

}

function is_ou_contain_manager ($OU) {
    # Проверить пристутствует ли руководитель в данной OU. Проверка идет по наличию сотрудников в AD-атрибуте "Прямые подчиненные".  

    if (Get-ADUser -SearchBase $OU  -SearchScope OneLevel -f {directReports -like "*" -and Enabled -eq "True"}) {
    
        return $true
       
    } else {

        return $false

    }

}

function add_user_to_group ($user, $OU) {
    # определяем в контейнер какого города закидывать пользователя и относительно этого выбираем дефолтную группу для пользователей элиты

    if ($OU.split(",")[-4] -eq "OU=company") {

       Add-ADGroupMember "CN=_Пользователи Программы,OU=company,DC=company,DC=net" -Member $user  

    } elseif ($OU.split(",")[-4] -eq "OU=KG") {

        Add-ADGroupMember "CN=KG_Users,OU=KGN,DC=company,DC=net" -Member $user 

    } elseif ($OU.split(",")[-4] -eq "OU=MC") {

        Add-ADGroupMember "CN=MC_Users,OU=MC,DC=company,DC=net" -Member $user   

    }
}

 $AD = [ordered]@{'1' = "OU=401,OU=company";
                  '2' = "OU=402,OU=company";
                  '3' = "OU=403,OU=company";
                  '4' = "OU=405,OU=company";
                  '5' = "OU=407,OU=company";
                  '6' = "OU=408,OU=company";}
                   

$chooseOuList = ("1 = 401",
                  "2 = 402",
                  "3 = 403",
                  "4 = 405",
                  "5 = 407",
                  "6 = 408")

#--------------------------------    Main Program    ---------------------------------#

$input = Read-Host "`nВведите логин или несколько через пробел" # ввод "lohin1 login2 login3..."
$rootDN = (Get-ADRootDSE).defaultNamingContext

foreach ($item in $input.split(" ")) {
    
    Write-Host "`n" 

    $info_hash = get_user_info($item)
    $info_hash.Keys | % { $info_hash.Item($_) } # выводим все что напарсили с сайта
    display_list $chooseOuList 4 8 # выводим список OU-шек указав кол-во столбцов и строк.

    $targetOU = Read-Host "`nВведите номер OU"
    $targetOU = "$($ad.Item($targetOU)),$rootDN" 
	
    if (is_manager $info_hash.title) {
        # Провереям является ли пользователь руководителем или начальником

        Write-Host "`nПользователь начальник или руководитель" -ForegroundColor Yellow -BackgroundColor DarkRed
        $manager = get_manager $targetOU # Руководитель в OU

        if ($manager -ne $null) {
        # Если руководитель присутствует в OU

            Write-Host "Предыдущий начальник в данной OU: " $manager.Name -ForegroundColor Yellow -BackgroundColor DarkRed
            $answer = Read-Host "`nЗаменить руководителя $($manager.Name) на $($info_hash.FIO) (Y/N)"

            if ($answer -eq "y") { 

                Write-Host "Заменить"
                change_manager $manager.DistinguishedName $info_hash.FIO $targetOU

            } Else {

                Write-Host "Не менять"

            }

        }
        
    }

    set_new_AD_user $info_hash $targetOU
    add_user_to_group $("$item") $targetOU
    #show_created_user_info $("test_$item")
    show_created_user_info $("$item")
}

Write-Host -NoNewLine 'Press any key to continue...';
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown'); 
exit


