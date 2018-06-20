#Get-Variable -Exclude PWD,*Preference | Remove-Variable -EA 0

########################################################################

$username = $env:USERNAME
$email = $username + 'company.net'
$appData = $env:APPDATA

$signPath = "$appdata\Microsoft\Signatures\sign.htm"
if (!(test-path "$appdata\Microsoft\Signatures")) {New-Item $appdata\Microsoft\Signatures -Type Directory}
# if (!(test-path $signPath)) {New-Item $appdata\Microsoft\Signatures -Type Directory} else { exit }
if (test-path $signPath) { exit }
#################### Generate HTML sign #################################

$userADProp = ([adsisearcher]"SamAccountName=$username").FindOne().properties
$manager = $userADProp.manager.Split(",")[0].substring(3)
$managerADProp = ([adsisearcher]"Name=$manager").FindOne().properties

$htmlSign = '<font size="2" color="grey" face="Calibri"><i>
	<br>С Уважением,
	<br> ' + $userADProp.displayname + ' 	
	<br> ' + $userADProp.title + '
	<br> ' + $userADProp.department + '
	<br> ' + $userADProp.company + '
	<br> тел.: ' + $userADProp.telephonenumber + ', ' + $userADProp.mobile + '
	<br> ' + $userADProp.mail + '
	<br> 
	<br> ' + $managerADProp.title + '
	<br> ' + $manager + '
	<br> ' + $managerADProp.telephonenumber + ' 
	<br> ' + $managerADProp.mail + '</i><br>
</font>' 


$htmlSign | out-File $signPath 











