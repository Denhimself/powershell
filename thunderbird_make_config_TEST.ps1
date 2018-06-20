Get-Variable -Exclude PWD,*Preference | Remove-Variable -EA 0   # for testing in poserwhell_ISE. Clear all variables.

$appData = $env:APPDATA
$username = $env:USERNAME
$profilePath = $appData + '\Thunderbird\';
$profileIni  = $profilePath + '\profiles.ini'
$newProfileName = $username + '@company.net';
$email = $username + '@company.net'

$profileIniContent ='
[General]
StartWithLastProfile=1

[Profile0]
Name=default
IsRelative=1
Path=Profiles/' + $newProfileName + '
Default=1
'


[string]$fullName = (([adsi]"WinNT://$env:userdomain/$username,user").fullname);
[string]$department = ([adsisearcher]"SamAccountName=$username").FindOne().properties.department
[string]$company = ([adsisearcher]"SamAccountName=$username").FindOne().properties.company
[string]$title = ([adsisearcher]"SamAccountName=$username").FindOne().properties.title
[string]$phoneNumber = ([adsisearcher]"SamAccountName=$username").FindOne().properties.telephonenumber

[string]$manager = ([adsisearcher]"SamAccountName=$username").FindOne().properties.manager
        $manager = $manager.Split(",")[0].substring(3)  
[string]$managerMail = ([adsisearcher]"Name=$manager").FindOne().properties.mail
[string]$managerTitle = ([adsisearcher]"Name=$manager").FindOne().properties.title
[string]$managerPhoneNumber = ([adsisearcher]"Name=$manager").FindOne().properties.telephonenumber



$sign =  '_________________________\nС Уважением,\n{0}\n{1}\n{2}\n{3}\nТел.: {8}\n{4}\n\n\n{5}\n{6}\nТел.: {9}\n{7}' -f `
         $fullname, $title, $department, $company, $email, $managerTitle, $manager, $managerMail, $phoneNumber, $managerPhoneNumber

$sign = $sign.Replace('"', '\"')
$company = $company.Replace('"', '\"')

$pref='
user_pref("mail.account.account1.identities", "id1");
user_pref("mail.account.account1.server", "server1");
user_pref("mail.account.account2.server", "server2");
user_pref("mail.account.lastKey", 2);
user_pref("mail.accountmanager.accounts", "account1,account2");
user_pref("mail.accountmanager.defaultaccount", "account1");
user_pref("mail.accountmanager.localfoldersserver", "server2");

user_pref("mail.append_preconfig_smtpservers.version", 2);

user_pref("mail.identity.id1.draft_folder", "mailbox://'+ $username +'@mail.company.net/Drafts");
user_pref("mail.identity.id1.drafts_folder_picker_mode", "0");
user_pref("mail.identity.id1.fcc_folder", "mailbox://'+ $username +'@mail.company.net/Sent");
user_pref("mail.identity.id1.fcc_folder_picker_mode", "0");
user_pref("mail.identity.id1.fullName", "'+ $fullName +'");

user_pref("mail.identity.id1.htmlSigText", "' + $sign + '");
user_pref("mail.identity.id1.organization", "'+ $company +'");
user_pref("mail.identity.id1.reply_to", "' + $managerMail + '");

user_pref("mail.identity.id1.reply_on_top", 1);
user_pref("mail.identity.id1.smtpServer", "smtp1");
user_pref("mail.identity.id1.stationery_folder", "mailbox://'+ $username +'@mail.company.net/Templates");
user_pref("mail.identity.id1.tmpl_folder_picker_mode", "0");
user_pref("mail.identity.id1.useremail", "'+ $email +'");
user_pref("mail.identity.id1.valid", true);

user_pref("mail.prompt_purge_threshhold", false);

user_pref("mail.root.none-rel", "[ProfD]Mail");
user_pref("mail.root.pop3-rel", "[ProfD]Mail");

user_pref("mail.server.server1.check_new_mail", true);
user_pref("mail.server.server1.check_time", 5);
user_pref("mail.server.server1.delete_by_age_from_server", true);
user_pref("mail.server.server1.delete_mail_left_on_server", false);
user_pref("mail.server.server1.directory-rel", "[ProfD]Mail/mail.company-1.net");
user_pref("mail.server.server1.download_on_biff", true);
user_pref("mail.server.server1.hostname", "mail.company.net");
user_pref("mail.server.server1.leave_on_server", true);
user_pref("mail.server.server1.login_at_startup", true);
user_pref("mail.server.server1.name", "'+ $email +'");
user_pref("mail.server.server1.num_days_to_leave_on_server", 14);
user_pref("mail.server.server1.socketType", 0);
user_pref("mail.server.server1.storeContractID", "@mozilla.org/msgstore/berkeleystore;1");
user_pref("mail.server.server1.type", "pop3");
user_pref("mail.server.server1.userName", "'+ $username +'");

user_pref("mail.smtpserver.smtp1.authMethod", 1);
user_pref("mail.smtpserver.smtp1.hostname", "mail.company.net");
user_pref("mail.smtpserver.smtp1.port", 25);
user_pref("mail.smtpserver.smtp1.try_ssl", 0);
user_pref("mail.smtpservers", "smtp1");

user_pref("mail.server.server1.num_days_to_leave_on_server", 4);
user_pref("mail.shell.checkDefaultClient", false);

user_pref("mailnews.start_page.enabled", false);
user_pref("network.proxy.type", 0);
'


if (-not(Test-Path $profilePath)) { # cheking existing of appdata\thunderbird

    New-Item -Path $profilePath\Profiles\$newProfileName -ItemType Directory; 
    New-Item -Path $profileIni -ItemType File
    Add-Content -Path $profileIni -Value $profileIniContent;
    Set-Content -Encoding utf8 -Path ($appData + '\Thunderbird\Profiles\' + $newProfileName + '\prefs.js') `
                -Value $pref 
    break;   
} 

$defaultProfileName = (Select-String '\w+\.(default)' -Path $profileIni) -split "/"; # get profile name folder
[string]$defaultProfileName = $defaultProfileName[1]

################################################################################

$currentDate = (get-date).ToString("dd.MM.yyyy")
$lastWriteDate = (Get-ChildItem $profileINI).LastWriteTime.ToString("dd.MM.yyyy")

if (Test-Path $profilePath\Profiles\$newProfileName) { <#"new profile name is exist"#> break; }   # checking the existence of a company profile directory
if ($lastWriteDate -ne $currentDate) { <#"date not equal"#> break; }

###############     Backup old profile     #####################################
Copy-Item ($profilePath + 'profiles.ini') ($profilePath + 'old_profiles.ini')

(get-content ($profilePath + 'profiles.ini')) -replace 'Path=Profiles/(.+.)',                  
("Path=Profiles/" + $newProfileName) | Set-Content ($profilePath + 'profiles.ini');

Copy-Item ($profilePath + '\Profiles\' + $defaultProfileName) `
          ($profilePath + '\Profiles\' + 'old_' + $defaultProfileName) -Recurse 
           
###############     Change profile directory name    ###########################
Rename-Item ($profilePath + '\Profiles\' + $defaultProfileName)`
            ($profilePath + '\Profiles\' + $newProfileName);

Start-Sleep -s 5
Remove-Item ($appData + '\Thunderbird\Profiles\' + $newProfileName + '\prefs.js')
Add-Content -Encoding utf8 -Path ($appData + '\Thunderbird\Profiles\' + $newProfileName + '\prefs.js') -Value $pref 