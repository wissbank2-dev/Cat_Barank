$WshShell = New-Object -ComObject WScript.Shell
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$Shortcut = $WshShell.CreateShortcut("$DesktopPath\Cat App.lnk")
$Shortcut.TargetPath = "C:\Create testcase test\Start_Cat_App.bat"
$Shortcut.WorkingDirectory = "C:\Create testcase test"
$Shortcut.Save()
Write-Host "Shortcut created on Desktop!"
