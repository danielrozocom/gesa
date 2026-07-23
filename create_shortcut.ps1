$WshShell = New-Object -ComObject WScript.Shell

# Start Menu Shortcut (Searchable in Windows Start)
$StartMenuPath = [System.IO.Path]::Combine($env:APPDATA, 'Microsoft\Windows\Start Menu\Programs\GESA.lnk')
$Shortcut = $WshShell.CreateShortcut($StartMenuPath)
$Shortcut.TargetPath = "$PSScriptRoot\Iniciar.bat"
$Shortcut.WorkingDirectory = "$PSScriptRoot"
$Shortcut.Description = "Gestor de Evaluaciones de Suficiencia Académica"
$Shortcut.Save()

# Desktop Shortcut
$DesktopPath = [System.IO.Path]::Combine($env:USERPROFILE, 'Desktop\GESA.lnk')
$ShortcutDesktop = $WshShell.CreateShortcut($DesktopPath)
$ShortcutDesktop.TargetPath = "$PSScriptRoot\Iniciar.bat"
$ShortcutDesktop.WorkingDirectory = "$PSScriptRoot"
$ShortcutDesktop.Description = "Gestor de Evaluaciones de Suficiencia Académica"
$ShortcutDesktop.Save()
