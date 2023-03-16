<#PSScriptInfo
.VERSION 1.4
.GUID 3b0fe24d-19b0-4ed7-a6ad-aab3e59f3b98
.AUTHOR daniil.zalivakhin@ma.ru
.COMPANYNAME Major Auto
.COPYRIGHT Major Auto
.DESCRIPTION - Скрипт для создания ярлыков подачи обращений в ИТ за поддержкой.
.RELEASENOTES
    2023-02-20: Изначальная версия.
#>
<#
.SYNOPSIS
    Скрипт создаёт ярлыки на рабочих столах пользователей для упрощения подачи заявок в ИТ за поддержкой. Ярлыки создаются на Windows 10 и 7.
   
.DESCRIPTION
   Скрипт создаёт ярлыки на рабочих столах пользователей. Ярлыки работают по типу ссылки "mailto:".
    При запуске ярлыка открывается письмо в outlook с заполненными полями:
        * Поле "Кому" - подставляется email IT SUPPORT, например "ITSUPPORT@info.local";
        * В тему письма подставляется имя ПК пользователя, например "ПК/IP: M47-WS1;".
    ВАЖНО! При первом запуске ярлыка, Windows может спросить через какое приложение открывать письмо - на выбор браузеры и Outlook, соответственно пользователь выбирает Outlook.
    Если пользователь выбрал приложением не Outlook, а браузер, то ярлык будет постоянно открывать выбранный браузер и работать не будет - нужно поменять настройки в Windows:
        - Приложения по умолчанию > Электронная почта > "Microsoft Outlook".
   
    Алгоритм работы скрипта:
        * Скрипт формирует список ПК для создания ярлыков и эскпортирует в файл CSV;
        * Работает по файлу, проверяет поле "LnkCreated", если значение "false" - создаёт ярлык и меняет значение на "true".


.NOTES
    Версия:          1.4 (2023-02-20)
    Автор:           Заливахин Даниил Антонович
    Должность:       системный администратор
#>

######### Ниже блок настройки скрипта #########
$OU_target = ""  #Здесь между кавычек указываем подразделение с ПК, на которых хотим создать ярлыки пользователям(Для windows 10 и windows 7);
$email_adress = ""       #Указываем эл. адрес IT SUPPORT *;
$lnk_name = "Заявка в IT"       #Название ярлыков;
######### Конец блока настройки скрипта #########

######################################################################################################################################
#[Console]::outputEncoding = [System.Text.Encoding]::GetEncoding('cp866')

Function Set-LnkDesktop {
    param ($PathToDesktop,$LnkName,$MailSubject,$IconArrayIndex)
    $IconLocation = "%SystemRoot%\system32\imageres.dll"
    $target = $PathToDesktop+"\$LnkName.lnk"
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($target)
    $Shortcut.TargetPath = $MailSubject
    $Shortcut.IconLocation = "$IconLocation, $IconArrayIndex"
    $Shortcut.Save()  
}

# Основная функция, вызывается ниже
Function Get-Main {
    param ($Path)
    $Source = (Get-Content $Path).count -2
    $PCListContent = Import-Csv -path $Path -Delimiter ';' -Encoding UTF8
    $Str = 0
    While($Str -notlike $Source) {
        # Проверяем создавался ли ярлык на ПК, для этого смотрим в таблице значение параметра "LnkCreated"
        if($PCListContent[$Str].LnkCreated -eq "False") {
            $PC_name = $PCListContent[$Str].Name
            # Если ПК онлайн, создаём ярлык на рабочем столе
            $check_connect = Test-Connection -ComputerName $PC_name -Quiet -Count 1 -ErrorAction SilentlyContinue
            if($check_connect) {
                Write-Host $PC_name "online"
                # Формируем тему письма, создаваемого ярлыком
                $MailSubj = 'mailto:'+$email_adress+'&subject=ПК/IP: '+$PC_name+';'
                # Поиск рабочих столов в профилях пользователей для публикации ярлыка
                $desktop_search = (Get-ChildItem "\\$PC_name\c$\Users\" -Recurse -Include "Desktop").FullName
                if($PCListContent[$Str].OperatingSystem -like "*10*") {
                    Write-Host $PC_name "Windows 10"
                    foreach($path_var in $desktop_search) {
                        Write-Host $path_var+"\$lnk_name.lnk"
                        Set-LnkDesktop -PathToDesktop $path_var -LnkName $lnk_name -MailSubject $MailSubj -IconArrayIndex 312 #create lnk with icon for windows 10  
                    }      
                } else {
                    if ($PCListContent[$Str].OperatingSystem -like "*7*") {
                        Write-Host $PC_name "Windows 7"
                        foreach($path_var in $desktop_search) {
                            Write-Host $path_var+"\$lnk_name.lnk"
                            Set-LnkDesktop -PathToDesktop $path_var -LnkName $lnk_name -MailSubject $MailSubj -IconArrayIndex 15 #create lnk with icon for windows 7
                        }
                    }
                }  
                # Отмечаем в CSV файле, что ярлык создался, меняем значение "True"
                $PCListContent[$Str].LnkCreated = "True"
                $PCListContent[$Str]
                $PCListContent | Export-Csv -Path $Path -Delimiter ';' -Encoding UTF8
            } else {
                # Пропускаем, ПК оффлайн
                Write-Host $PCListContent[$Str].Name "offline"
            }
        }  
        # Повышаем итерацию для цикла и для перебора строк в файле
        $Str++
    } # Конец цикла "While($Str -notlike $Source)"
    $PCListContent
} # Конец функции "Get-Main"

# Указываем расположение файла со списком ПК
$PCListPath = "$env:USERPROFILE\AppData\Local\Temp\PC_list.csv"
# Если файл со списком ПК создан - продолжаем работать по нему
if([System.IO.File]::Exists("$PCListPath")) {
    Write-Host "File Exists"
    Get-Main $PCListPath
} else {
    # Иначе формируем список ПК, создаём файл CSV
    Write-Host "File not Exists"
    # Делаем один LDAP запрос на контроллер домена
    Get-ADComputer -Properties Name,OperatingSystem,Enabled -Filter {Enabled -eq $true} -SearchBase $OU_target | Sort-Object OperatingSystem |
    Select-Object Name,OperatingSystem, @{Name="LnkCreated";Expression={"False"}} |
    #Select-Object Name,OperatingSystem, @{Name="Created";Expression={$_.Description = "-"}} |
    Export-CSV $PCListPath -Delimiter ';' -Encoding UTF8 #-NoTypeInformation
    Get-Main $PCListPath
}