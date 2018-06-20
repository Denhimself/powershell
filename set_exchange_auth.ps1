Get-Variable -Exclude PWD,*Preference | Remove-Variable -EA 0

function encodingPsiPassword ($jid, $password) {
    
    $encodePass = $password  
    [System.Collections.ArrayList]$decodePass = @{}

    ########  Adding to array Each four symbols from encoding password and convert to DEC  ############

    for($i=0; $i -lt $encodePass.Length; $i += 4) {

        #$decodePass.Add([Convert]::ToInt32($encodePass.Substring($i, 4),16))
        $dec = [Convert]::ToInt32($encodePass.Substring($i, 4),16)
        $decodePass += $dec

    }


    #######   Decoding   #######################

    $i = 0
    $x = 0

    [string]$result = '';
    #$jid = 'vdi@jabber.net'

    foreach ($char in $decodePass) {

        $x = $char -bxor ([char]($jid[$i % ($jid.Length)]))
        $i += 1;
        $result += [char]$x

    }

    return $result

}


function get_jid_and_pass {

    $homedir =  $env:USERPROFILE 
    $configPath = "$homeDir\PsiData\profiles\default\config.xml"

    if (Test-Path $configPath) {
    
        $psiUserInfo = Select-Xml -Path $configPath  -XPath "//account" | Select -ExpandProperty node | Select jid, password

    } else { break }


    return $psiUserInfo

}


function set_mail_authentication ($password) {

    $lanLogin = "company.lan\$env:USERNAME"
    #$password = $encodingPass
    $authenticationList = ("*.company.lan", "x.company.com", "autodiscover.tu.company.com", "autodiscover.company.com")

    foreach ($address in $authenticationList) {
   
        cmdkey /add:$address /U:$lanLogin /pass:$password

    }

}

###########################   Main program    ############################

$psiUserInfo = get_jid_and_pass

[string]$jid = $psiUserInfo.jid
[string]$password = $psiUserInfo.password

$encodingPass = encodingPsiPassword $jid.Trim() $password.Trim()
set_mail_authentication $encodingPass


