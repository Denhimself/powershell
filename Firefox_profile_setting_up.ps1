###############    add settings and IE_tab_plugin    #####################
    Function settingUp {

    
        Add-content -Path ($profilePath + '\Profiles\' + $newProfileName + '\prefs.js') `
                    -Value $prefsJS; 

        Set-Content -Path ($profilePath + '\Profiles\' + $newProfileName + '\extensions.ini') `
                    -Value $extensionsINI;

        Copy-Item -Path ('.\{1BC9BA34-1EED-42ca-A505-6D2F1A935BBB}\') -Recurse `
                  -Destination ($profilePath + '\Profiles\' + $newProfileName + '\extensions\{1BC9BA34-1EED-42ca-A505-6D2F1A935BBB}');
    } 

#############      Firefox Settings         ##############################

$prefsJS = ' 

user_pref("browser.startup.homepage", "komus.net");
user_pref("browser.search.update", false);

user_pref("browser.cache.disk.capacity", 102400);
user_pref("browser.cache.disk.smart_size.enabled", false);
user_pref("browser.cache.disk.smart_size.first_run", false);
user_pref("browser.cache.disk.smart_size.use_old_max", false);

user_pref("browser.startup.homepage_override.buildID", " ");
user_pref("browser.startup.homepage_override.mstone", " ");

user_pref("dom.disable_open_during_load", false);

user_pref("app.update.auto", false);
user_pref("app.update.enabled", false);
user_pref("app.update.service.enabled", false);
user_pref("datareporting.healthreport.uploadEnabled", false);

user_pref("xpinstall.signatures.required", false);

user_pref("extensions.ietab2.ietab2PrefsMigrated", true);
user_pref("extensions.ietab2.prefsMigrated", true);
user_pref("extensions.ietab2.filterlist", "*sdo.komus.net* *sed.komus.net* /^file:\\/\\/\\/.*\\.(mht|mhtml)$/ http://*update.microsoft.com/* http://www.windowsupdate.com/*");

user_pref("extensions.update.autoUpdateDefault", false);
' +
#user_pref("extensions.e10sBlockedByAddons", false);
#user_pref("extensions.enabledAddons", "%7B1BC9BA34-1EED-42ca-A505-6D2F1A935BBB%7D:4.1.3.1,%7B1BC9BA34-1EED-42ca-A505-6D2F1A935BBB%7D:4.1.3.1");
'
user_pref("network.proxy.type", 5);

';
##########################################################################
   
       $profilePath = $env:APPDATA + '\Mozilla\Firefox\';
       $profileIni  = $profilePath + '\profiles.ini'
    $newProfileName = $env:USERNAME + '@company.net';
    $profileIniContent ='
[General]
StartWithLastProfile=1

[Profile0]
Name=default
IsRelative=1
Path=Profiles/' + $newProfileName + '
Default=1
';
$extensionsINI =   
'
[ExtensionDirs]
Extension0=' + $profilePath + 'Profiles\' + $newProfileName + '\extensions\{1BC9BA34-1EED-42ca-A505-6D2F1A935BBB}

[MultiprocessIncompatibleExtensions]
Extension0={1BC9BA34-1EED-42ca-A505-6D2F1A935BBB}

';

if (-not(Test-Path $profilePath)) { # проверка существования appdata\mozilla

    New-Item -Path $profilePath\Profiles\$newProfileName -ItemType Directory; 
    New-Item -Path $profileIni -ItemType File
    Add-Content -Path $profileIni -Value $profileIniContent;
    settingUp;
    break;   
} 

$defaultProfileName = (Select-String '\w+\.(default)' -Path $profileIni) -split "/";


############################################################################

$currentDate = (get-date).ToString("dd.MM.yyyy")
$lastWriteDate = (Get-ChildItem $profileINI).LastWriteTime.ToString("dd.MM.yyyy")

if (Test-Path $profilePath\Profiles\$newProfileName) { <#"new profile name is exist"#> break; }   # проверка наличия профиля созданного скриптом
if ($lastWriteDate -ne $currentDate) { <#"date not equal"#> break; } # проверка даты создания профиля

###############     Backup old profile     #############################
Copy-Item ($profilePath + 'profiles.ini') ($profilePath + 'old_profiles.ini')

(get-content ($profilePath + 'profiles.ini')) -replace 'Path=Profiles/(.+.)',                  
("Path=Profiles/" + $newProfileName) | Set-Content ($profilePath + 'profiles.ini');

Copy-Item ($profilePath + '\Profiles\' + $defaultProfileName[1]) `
          ($profilePath + '\Profiles\' + 'old_' + $defaultProfileName[1]) 
           
###############     Change profile directory name    #####################
Rename-Item ($profilePath + '\Profiles\' + $defaultProfileName[1])`
            ($profilePath + '\Profiles\' + $newProfileName);


SettingUp;

