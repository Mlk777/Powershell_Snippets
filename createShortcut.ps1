<#
New shortcut as administrator
Change "TargetPath" depending on the shortcut (cmd.exe) or (cscript.exe)
#>

$filePath = "" #Source file path
$icon = "" #Ico file 
$shortcut = ""#Shortcut path + shortcut name

$WshShell = New-Object -ComObject WScript.shell
$shortcut = $WshShell.CreateShortcut("$shortcut")
$shortcut.TargetPath = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe' 
$shortcut.arguments = '-file ' + "$filePath"
$shortcut.IconLocation = "$icon"
$shortcut.save()

$bytes = [System.IO.File]::ReadAllBytes("$shortcut")
$bytes[0x15] = $bytes[0x15] -bor 0x20 
[System.IO.File]::WriteAllBytes("$shortcut", $bytes)
