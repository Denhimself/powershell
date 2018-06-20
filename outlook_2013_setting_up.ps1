If (-not (test-path HKLM:SOFTWARE\Classes\outlook.Application)) {
    #write-host "outlook is not installed"
	break
}

if (Test-Path HKCU:Software\Microsoft\Office\15.0\Outlook\Profiles\*) {
    #write-host "outlook profile is exist"
    break
} 


$userprofile = $env:USERPROFILE
$user = $env:USERNAME
$email = $env:USERNAME + '@company.net';
$appData = $env:APPDATA
$fullName = (([adsi]"WinNT://$env:userdomain/$env:username,user").fullname) 
$prf ='
;Automatically generated PRF file from the Microsoft Office Customization and Installation Wizard

; **************************************************************
; Section 1 - Profile Defaults
; **************************************************************

[General]
Custom=1
ProfileName=' + $user + '@company.net
DefaultProfile=No
OverwriteProfile=No
ModifyDefaultProfileIfPresent=false
BackupProfile=Yes
;DefaultStore=' + $user + '@company.net
DefaultStore=Service2

; **************************************************************
; Section 2 - Services in Profile
; **************************************************************

[Service List]
;ServiceX=Microsoft Outlook Client
;Service1=LDAP Directory
Service2=Unicode Personal Folders
;Service3=Personal Folders

;***************************************************************
; Section 3 - List of internet accounts
;***************************************************************

[Internet Account List]
Account1=I_Mail

;***************************************************************
; Section 4 - Default values for each service.
;***************************************************************


[Service1]
UniqueService=No
ServerName=company.net
DisplayName=company.net
ConnectionPort=389
UseSSL=false
UseSPA=false
EnableBrowsing=true
UserName=' + $user + '@company.net
SearchBase=DC=company74,DC=company,DC=net
DefaultSearch=yes
SearchTimeout=60
MaxEntriesReturned=100
;CheckNames=

[Service2]
UniqueService=No
Name=' + $email  + '
PathToPersonalFolders=%USERPROFILE%\documents\Файлы Outlook\' + $email  + '.pst
EncryptionType=0x50000000


;***************************************************************
; Section 5 - Values for each internet account.
;***************************************************************

[Account1]
UniqueService=No
AccountName=' + $email +'
POP3Server=mail.company.net
SMTPServer=mail.company.net
POP3UserName=' + $user + '
EmailAddress=' + $email +'
POP3UseSPA=0
DisplayName=' + $fullName +' 
Organization=ООО "Мелкомягкий"
ReplyEMailAddress=' + $email +'
SMTPUseAuth=0
SMTPAuthMethod=0
ConnectionType=0
LeaveOnServer=0x20007
POP3UseSSL=0
ConnectionOID=MyConnection
POP3Port=110
ServerTimeOut=60
SMTPPort=25
SMTPSecureConnection=0

;***************************************************************
; Section 6 - Mapping for profile properties
;***************************************************************

    [Microsoft Exchange Server]
    

    [Exchange Global Section]
    

    [Microsoft Mail]
    

    [Personal Folders]
  

    [Unicode Personal Folders]
    ServiceName=MSPST MS
    Name=PT_STRING8,0x3001
    PathToPersonalFolders=PT_STRING8,0x6700 
    RememberPassword=PT_BOOLEAN,0x6701
    EncryptionType=PT_LONG,0x6702
    Password=PT_STRING8,0x6703

    [Outlook Address Book]
    ServiceName=CONTAB

    [LDAP Directory]
    ServiceName=EMABLT
    ServerName=PT_STRING8,0x6600
    UserName=PT_STRING8,0x6602
    UseSSL=PT_BOOLEAN,0x6613
    UseSPA=PT_BOOLEAN,0x6615
    EnableBrowsing=PT_BOOLEAN,0x6622
    DisplayName=PT_STRING8,0x3001
    ConnectionPort=PT_STRING8,0x6601
    SearchTimeout=PT_STRING8,0x6607
    MaxEntriesReturned=PT_STRING8,0x6608
    SearchBase=PT_STRING8,0x6603
    CheckNames=PT_STRING8,0x6624
    DefaultSearch=PT_LONG,0x6623

    [Microsoft Outlook Client]
    

    [Personal Address Book]
    ServiceName=MSPST AB
    NameOfPAB=PT_STRING8,0x001e3001
    PathAndFilename=PT_STRING8,0x001e6600
    ShowNamesBy=PT_LONG,0x00036601

; ************************************************************************
; Section 7 - Mapping for internet account properties.  DO NOT MODIFY.
; ************************************************************************

[I_Mail]
AccountType=POP3
--- POP3 Account Settings ---
AccountName=PT_UNICODE,0x0002
DisplayName=PT_UNICODE,0x000B
EmailAddress=PT_UNICODE,0x000C
--- POP3 Account Settings ---
POP3Server=PT_UNICODE,0x0100
POP3UserName=PT_UNICODE,0x0101
POP3UseSPA=PT_LONG,0x0108
Organization=PT_UNICODE,0x0107
ReplyEmailAddress=PT_UNICODE,0x0103
POP3Port=PT_LONG,0x0104
POP3UseSSL=PT_LONG,0x0105
 --- SMTP Account Settings ---
SMTPServer=PT_UNICODE,0x0200
SMTPUseAuth=PT_LONG,0x0203
SMTPAuthMethod=PT_LONG,0x0208
SMTPUserName=PT_UNICODE,0x0204
SMTPUseSPA=PT_LONG,0x0207
ConnectionType=PT_LONG,0x000F
ConnectionOID=PT_UNICODE,0x0010
SMTPPort=PT_LONG,0x0201
SMTPSecureConnection=PT_LONG,0x020A
ServerTimeOut=PT_LONG,0x0209
LeaveOnServer=PT_LONG,0x1000

    [IMAP_I_Mail]
    
'
##############################################################################################################################################################

Copy-Item ('.\company_default.pst')($userprofile + '\Documents\Файлы Outlook\' + $email  + '.pst') # Copy and rename empty pst file 
Add-Content -Path ($appdata + '\outlook_company.prf') -Value $prf; # Create and fill a prf file
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\15.0\Common\General' -Name 'ShownFirstRunOptin' -Value '1' # disable first run option panel

Start-Process "C:\Program Files (x86)\Microsoft Office\Office15\OUTLOOK.EXE" -ArgumentList ("/nopreview " + "/importprf " + $appdata + "\outlook_company.prf") -NoNewWindow -Wait; 


###############################################################################################################################################################

$registryPath = "HKCU:\Software\Microsoft\Office\15.0\Outlook\Profiles\" + $email + "\"
$Name = "0003041b"
$keyName = (Get-ItemProperty -Path $registryPath\* -Name "0003041b" -ErrorAction SilentlyContinue | Select-Object pschildname | Format-Wide -Property pschildname) | out-string

$keyName = $keyName.Split() | where {$_}
$fullPath = $registryPath+$keyName

$securityLevel = "ff,ff,ff,ff"
$hexified = $securityLevel.Split(',') | % { "0x$_"}
Set-ItemProperty -Path $fullPath -Name $name -Value ([byte[]]$hexified) 

$name="001f0418"
$trustedSenders="40,0,6B,0,6F,0,6D,0,75,0,73,0,2E,0,6E,0,65,0,74,0,3B,0,40,0,72,0,65,0,67,0,69,0,6F,0,6E,0,2E,0,6B,0,6F,0,6D,0,75,0,73,0,2E,0,6E,0,65,0,74,0,3B,0,0,0"
$hexified = $trustedSenders.Split(',') | % { "0x$_"}
Set-ItemProperty -Path $fullPath -Name $name -Value ([byte[]]$hexified)


Remove-Item -Path ($appdata + '\outlook_company.prf')

#################################### Setting up a send\recieve time = 2 minuties ##############################################################################

$srsFileName = (Get-itemproperty -path ("HKCU:\Software\Microsoft\Office\15.0\Outlook\Profiles\" + $email + "\3517490d76624c419a828607e2a54604"))."001f6000" 
[string]$srsFileName = $srsFileName | %{[CHAR][BYTE]$_} 
$srsFileName = ($srsFileName -replace " `0 ", "").TrimEnd("`0", " ") + ".srs"

[string]$srsContent = (Get-content -path ($appdata + "\Microsoft\Outlook\" + $srsFileName))
$srsContentArray = @($srsContent.ToCharArray())
$srsContentArray[2163] = [char]2
$srsContentArray[2169] = [char]2


Set-Content -Path ($appdata + "\Microsoft\Outlook\" + $srsFileName) -Value (-join $srsContentArray)

#################################################################################################################################################################

