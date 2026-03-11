$DesktopPath = [System.Environment]::GetFolderPath('Desktop')
$ShortcutPath = "$DesktopPath\ReportStock.lnk"
$BatchFile = "$DesktopPath\ReportStock.bat"
$IconPath = "$DesktopPath\icon.ico"

$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortCut($ShortcutPath)
$Shortcut.TargetPath = "cmd.exe"
$Shortcut.Arguments = "/c `"$BatchFile`""
$Shortcut.WorkingDirectory = "$DesktopPath"
$Shortcut.IconLocation = $IconPath
$Shortcut.WindowStyle = 1
$Shortcut.Save()

Write-Host "Atajo creado: $ShortcutPath"
