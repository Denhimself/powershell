#Get-Variable -Exclude PWD,*Preference | Remove-Variable -EA 0  #удалить строку при звпуске через политику
 
######################################################################################################

##################################    Password Encoding.     #########################################

###################################################################################################### 
                                                              
function encodingPsiPassword ($jid, $password )
{
    [string]$result = '';
    $x = 0;
    $i = 0;

    foreach ($char in ($password.GetEnumerator())) {
        $x = $char -bxor ([int32][char]($jid[$i % ($jid.Length)]))
        $result += '{0:X4}' -f $x 
        $i += 1;
    }

    return $result
}

######################################################################################################

   $currentUser = $env:USERNAME;
           $jid = $currentUser + '@jabber.net'
       $homedir =  $env:USERPROFILE                                                     

######################################################################################################

if (!(Test-Path "C:\Program Files (x86)\Psi\psi.exe")) 
    {
        if (!(Test-Path "C:\Program Files\Psi\psi.exe")) {break} 
    }

if (!(Test-Path $homeDir\PsiData\profiles\default\config.xml))              
    {   
        if (!(Test-Path $homeDir\PsiData\profiles\default))
            {                                                                   # проверка существования профиля 
                New-Item $homeDir\PsiData\profiles\default -ItemType Directory    
            }
    }
else { break }

if (Get-Process -Name "psi" -ErrorAction SilentlyContinue) {
    Stop-Process -Name psi -Force | Out-Null                           # завершить процесс Psi 
}

Clear-Host

$password = Read-Host 'Введите пароль ЕКА и нажмите Enter '
$EcodedPassword = encodingPsiPassword $jid $password

######################################################################################################

##################   Содержимое Config.xml для jabber-client Psi.     ###############################

######################################################################################################

$xml = '<psiconf version="1.0" >
        <progver>0.9.3</progver>
         <geom>1350,147,244,570</geom>
         <recentGCList/>
         <recentBrowseList>
          <item>jabber1.komus.net</item>
         </recentBrowseList>
         <lastStatusString></lastStatusString>
         <useSound>true</useSound>
         <accounts>
         <account showSelf="false" plain="true" ssl="false" ignoreSSLWarnings="false" showHidden="true" keepAlive="true" showAgents="true" showOffline="true" showAway="true" log="true" auto="true" reconn="true" enabled="true" >
           <name>'+$currentUser+'</name>
           <jid>'+$jid+'</jid> 
           <password>'+$EcodedPassword+'</password>
           <useHost>false</useHost>
           <host></host>
           <port>5222</port>
           <resource>Psi</resource>
           <priority>5</priority>
           <OLR></OLR>
           <pgpSecretKeyID></pgpSecretKeyID>
           <groupState>
            <group open="true" name="/\/'+$currentUser+'\/\" rank="0" />
           </groupState>
           <proxyindex>0</proxyindex>
           <pgpkeybindings/>
           <dtProxy></dtProxy>
          </account>
         </accounts>
         <proxies/>
         <preferences>
          <general>
           <roster>
            <useleft>false</useleft>
            <singleclick>false</singleclick>
            <defaultAction>1</defaultAction>
            <useTransportIconsForContacts>false</useTransportIconsForContacts>
            <sortStyle>
             <contact>status</contact>
             <group>alpha</group>
             <account>alpha</account>
            </sortStyle>
           </roster>
           <misc>
            <smallChats>false</smallChats>
            <delChats>1</delChats>
            <showJoins>true</showJoins>
            <browser>0</browser>
            <customBrowser></customBrowser>
            <customMailer></customMailer>
            <alwaysOnTop>true</alwaysOnTop>
            <keepSizes>true</keepSizes>
            <ignoreHeadline>false</ignoreHeadline>
            <ignoreNonRoster>false</ignoreNonRoster>
            <excludeGroupChatIgnore>true</excludeGroupChatIgnore>
            <scrollTo>true</scrollTo>
            <useEmoticons>false</useEmoticons>
            <alertOpenChats>false</alertOpenChats>
            <raiseChatWindow>false</raiseChatWindow>
            <showSubjects>true</showSubjects>
            <showCounter>false</showCounter>
            <chatSays>false</chatSays>
            <chatSoftReturn>true</chatSoftReturn>
            <showGroupCounts>true</showGroupCounts>
            <jidComplete>true</jidComplete>
            <grabUrls>false</grabUrls>
            <messageEvents>true</messageEvents>
           </misc>
           <dock>
            <useDock>true</useDock>
            <dockDCstyle>true</dockDCstyle>
            <dockHideMW>false</dockHideMW>
            <dockToolMW>false</dockToolMW>
            <isWMDock>false</isWMDock>
           </dock>
          </general>
          <ssl>
           <trustedcertstoredir>C:/Program Files (x86)/Psi/certs</trustedcertstoredir>
          </ssl>
          <events>
           <alertstyle>2</alertstyle>
           <autoAuth>false</autoAuth>
           <receive>
            <popupMsgs>false</popupMsgs>
            <popupChats>false</popupChats>
            <popupHeadlines>false</popupHeadlines>
            <popupFiles>false</popupFiles>
            <noAwayPopup>false</noAwayPopup>
            <noUnlistedPopup>false</noUnlistedPopup>
            <raise>false</raise>
            <incomingAs>0</incomingAs>
           </receive>
           <priority>
            <message>1</message>
            <chat>1</chat>
            <headline>0</headline>
            <auth>2</auth>
            <file>3</file>
           </priority>
          </events>
          <presence>
           <misc>
            <askOnline>false</askOnline>
            <askOffline>false</askOffline>
            <rosterAnim>true</rosterAnim>
            <autoVersion>false</autoVersion>
            <autoVCardOnLogin>true</autoVCardOnLogin>
            <xmlConsoleOnLogin>false</xmlConsoleOnLogin>
           </misc>
           <autostatus>
            <away use="true" >10</away>
            <xa use="true" >30</xa>
            <offline use="false" >0</offline>
            <message>Автостатус (неактивен)</message>
           </autostatus>
           <statuspresets>
            <item name="Нет на месте" >Меня нет на месте. Оставьте сообщение.</item>
            <item name="Принимаю душ" >Я в душе.  Вам придется подождать.</item>
            <item name="Ем" >Отошел поесть.  Ммм... еду.</item>
            <item name="Сплю" >Сон нормальный.  Хр-р-р</item>
            <item name="Работаю" >Некогда болтать.  Надо работать.</item>
            <item name="Гуляю" >Вышел прогуляться на свежем воздухе.</item>
            <item name="В кино" >Пошел в кино.  Может вместе?</item>
            <item name="Секрет" >Меня нет на месте и это все, что вам положено знать.</item>
            <item name="Дома" >Пошел на ночь домой.</item>
            <item name="В Греции" >Я очень далеко.  Но однажды я вернусь!</item>
           </statuspresets>
           <recentstatus/>
          </presence>
          <lookandfeel>
           <newHeadings>true</newHeadings>
           <colors>
            <online>#0060c0</online>
            <listback>#c0c0c0</listback>
            <away>#008080</away>
            <dnd>#800000</dnd>
            <offline>#000000</offline>
            <groupfore>#000000</groupfore>
            <groupback>#ffffff</groupback>
            <animfront>#ff0000</animfront>
            <animback>#000000</animback>
           </colors>
           <fonts>
            <roster>MS Shell Dlg,8.3,-1,5,50,0,0,0,0,0</roster>
            <message>MS Shell Dlg,8.3,-1,5,50,0,0,0,0,0</message>
            <chat>MS Shell Dlg,8.3,-1,5,50,0,0,0,0,0</chat>
            <popup>MS Shell Dlg,6,-1,5,50,0,0,0,0,0</popup>
           </fonts>
          </lookandfeel>
          <sound>
           <player>play</player>
           <noawaysound>false</noawaysound>
           <noGCSound>true</noGCSound>
           <onevent>
            <message>C:/Program Files (x86)/Psi/sound/chat2.wav</message>
            <chat1>C:/Program Files (x86)/Psi/sound/chat1.wav</chat1>
            <chat2>C:/Program Files (x86)/Psi/sound/chat2.wav</chat2>
            <system>C:/Program Files (x86)/Psi/sound/chat2.wav</system>
            <headline>C:/Program Files (x86)/Psi/sound/chat2.wav</headline>
            <online>C:/Program Files (x86)/Psi/sound/online.wav</online>
            <offline>C:/Program Files (x86)/Psi/sound/offline.wav</offline>
            <send>C:/Program Files (x86)/Psi/sound/send.wav</send>
            <incoming_ft>C:/Program Files (x86)/Psi/sound/ft_incoming.wav</incoming_ft>
            <ft_complete>C:/Program Files (x86)/Psi/sound/ft_complete.wav</ft_complete>
           </onevent>
          </sound>
          <sizes>
           <eventdlg>880,924</eventdlg>
           <chatdlg>790,420</chatdlg>
          </sizes>
          <toolbars>
           <mainWin>
            <toolbar>
             <name>Кнопки</name>
             <on>true</on>
             <locked>true</locked>
             <stretchable>true</stretchable>
             <keys>
              <item>button_options</item>
              <item>button_status</item>
             </keys>
             <position>
              <dock>DockBottom</dock>
              <index>1</index>
              <nl>true</nl>
              <extraOffset>0</extraOffset>
             </position>
            </toolbar>
            <toolbar>
             <name>Показать контакты</name>
             <on>true</on>
             <locked>true</locked>
             <stretchable>false</stretchable>
             <keys>
              <item>show_offline</item>
              <item>show_away</item>
              <item>show_agents</item>
              <item>show_self</item>
             </keys>
             <position>
              <dock>DockTop</dock>
              <index>2</index>
              <nl>true</nl>
              <extraOffset>0</extraOffset>
             </position>
            </toolbar>
            <toolbar>
             <name>Оповещатель о событиях</name>
             <on>false</on>
             <locked>true</locked>
             <stretchable>true</stretchable>
             <keys>
              <item>event_notifier</item>
             </keys>
             <position>
              <dock>DockBottom</dock>
              <index>0</index>
              <nl>true</nl>
              <extraOffset>0</extraOffset>
             </position>
            </toolbar>
           </mainWin>
          </toolbars>
          <popups>
           <on>false</on>
           <online>true</online>
           <offline>true</offline>
           <statusChange>false</statusChange>
           <message>true</message>
           <chat>true</chat>
           <headline>true</headline>
           <file>true</file>
           <jidClip>25</jidClip>
           <statusClip>-1</statusClip>
           <textClip>300</textClip>
           <hideTime>10000</hideTime>
           <borderColor>#5297f9</borderColor>
          </popups>
          <groupchat>
           <highlightwords/>
           <nickcolors>
            <item>Blue</item>
            <item>Green</item>
            <item>Orange</item>
            <item>Purple</item>
            <item>Red</item>
           </nickcolors>
           <nickcoloring>true</nickcoloring>
           <highlighting>true</highlighting>
          </groupchat>
          <lockdown>
           <roster>false</roster>
           <options>false</options>
           <profiles>false</profiles>
           <services>false</services>
           <accounts>false</accounts>
          </lockdown>
          <iconset>
           <system>default</system>
           <roster>
            <default>default</default>
            <service>
             <item service="aim" iconset="aim" />
             <item service="gadugadu" iconset="gadugadu" />
             <item service="icq" iconset="icq" />
             <item service="msn" iconset="msn" />
             <item service="sms" iconset="sms" />
             <item service="transport" iconset="transport" />
             <item service="yahoo" iconset="yahoo" />
            </service>
            <custom/>
           </roster>
           <emoticons>
            <item>default</item>
           </emoticons>
          </iconset>
          <tipOfTheDay>
           <show>false</show>
           <num>1</num>
          </tipOfTheDay>
          <disco>
           <items>true</items>
           <info>true</info>
          </disco>
          <dt>
           <port>8010</port>
           <external></external>
          </dt>
         </preferences>
        </psiconf>'

####################################################################################################################

$xml | Set-Content  $homeDir\PsiData\profiles\default\config.xml -Encoding utf8 # Создать конфигурационный файл 

if (Test-Path "C:\Program Files (x86)\Psi\psi.exe") 
    {
        Start-Process "C:\Program Files (x86)\Psi\psi.exe"
    }

if (Test-Path "C:\Program Files\Psi\psi.exe") 
    {
        Start-Process "C:\Program Files\Psi\psi.exe"
    }