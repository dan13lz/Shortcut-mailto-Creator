$OU_target = "OU=TEST,OU=IT,OU=M47,OU=MAJOR,DC=aamajor,DC=local"  #Здесь между кавычек указываем подразделение с ПК, на которых хотим создать ярлыки пользователям(Для windows 10 и windows 7);
$email_adress = "itsupportmkad@info.local"       #Указываем наш эл. адрес;
$lnk_name = "Заявка в IT MKAD"       #Название ярлыков;

######################################################################################################################################
[Console]::outputEncoding = [System.Text.Encoding]::GetEncoding('cp866')
Function Write-HostAndLog {                        
    param ($FuncWHLText,$FuncWHLOutFile)
    Write-Host $FuncWHLText             
    Add-Content $FuncWHLOutFile $FuncWHLText        # Get-Content $path_log_Found | Select-Object -Unique
   }
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

#<Create_log>
Start-Transcript -Path "$user\AppData\Local\Temp\lnk_creator_logs\PS_Transcript.log" -Force
$user = $env:USERPROFILE
$date = Get-Date -Format "HH:mm:ss dd/MM/yyyy" | ForEach-Object { $_ -replace ":", "-" }
$path_log_NotFound = "$user\AppData\Local\Temp\lnk_creator_logs\NotReachableHosts_$date.log" <#-----------offline hosts list-----------#>
$path_log_Found = "$user\AppData\Local\Temp\lnk_creator_logs\LnkEstablished.log"    <#-----------lnk created list-----------#>
new-item -path $path_log_NotFound -force   
$path_nextHop = "$user\AppData\Local\Temp\lnk_creator_logs\!next_hop.log"
#</Create_log>
if([System.IO.File]::Exists("$path_nextHop")){
    $PC_list = get-content $path_nextHop
    $PC_list = get-content $PC_list
}
else{
    $PC_list = get-adcomputer -searchbase $OU_target -Filter * | Select-object -ExpandProperty name
}
$PC_count = ($PC_list).count
$line_number = 0          
$PC_name = 0        
While ($line_number -notlike $PC_count){
    $PC_name = $PC_list[$line_number]   
    $source = 'mailto:'+$email_adress+'&subject=ПК/IP: '+$PC_name+';'
    $check_connect = Test-Connection -ComputerName $PC_name -Quiet -Count 1 -ErrorAction SilentlyContinue
    if($check_connect -and (Get-ADComputer -Identity $PC_name -Property * | Select-object -ExpandProperty Enabled)){
        $desktop_search = (Get-ChildItem "\\$PC_name\c$\Users\" -Recurse -Include "Desktop").FullName 
        $check_os = Get-ADComputer -Identity $PC_name -Property * | Select-object -ExpandProperty OperatingSystem  
        if($check_os -like '*10*'){
            write-host(".lnk on "+$PC_name+" creating...")
            foreach($path_var in $desktop_search){       
                Set-LnkDesktop -PathToDesktop $path_var -LnkName $lnk_name -MailSubject $source -IconArrayIndex 312   #create lnk with icon for windows 10
            } 
            $write_log = "("+$date+") "+$PC_name+" lnk created - Windows 10"
            Write-HostAndLog "$write_log" "$path_log_Found"   
        }   
        else { #-like windows 7 creating:
            write-host(".lnk on "+$PC_name+" creating...")
            foreach($path_var in $desktop_search){       
                Set-LnkDesktop -PathToDesktop $path_var -LnkName $lnk_name -MailSubject $source -IconArrayIndex 15   #create lnk with icon for windows 7
            } 
            $write_log = "("+$date+") "+$PC_name+" lnk created - Windows 7"            
            Write-HostAndLog "$write_log" "$path_log_Found"    
        } #end check_os
    } 
    else{
            Write-HostAndLog "$PC_name" "$path_log_NotFound"              
    } #end check_connect
    $line_number++
} #end while
set-Content $path_nextHop $path_log_NotFound
[Console]::outputEncoding 
Stop-Transcript
#notepad $path_log_Found
#notepad "$user\AppData\Local\Temp\lnk_creator_logs\PS_Transcript.log"