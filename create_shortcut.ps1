$WshShell = New-Object -ComObject WScript.Shell
$ExePath = [System.IO.Path]::Combine($PSScriptRoot, 'GESA.exe')
$IconPath = [System.IO.Path]::Combine($PSScriptRoot, 'app_icon.ico')

$TargetPath = if (Test-Path $ExePath) { $ExePath } else { [System.IO.Path]::Combine($PSScriptRoot, 'start.bat') }

# Start Menu Shortcut (Searchable in Windows Start)
$StartMenuPath = [System.IO.Path]::Combine($env:APPDATA, 'Microsoft\Windows\Start Menu\Programs\GESA.lnk')
$Shortcut = $WshShell.CreateShortcut($StartMenuPath)
$Shortcut.TargetPath = $TargetPath
$Shortcut.WorkingDirectory = "$PSScriptRoot"
$Shortcut.Description = "Gestor de Evaluaciones de Suficiencia Académica"
if (Test-Path $IconPath) {
    $Shortcut.IconLocation = "$IconPath,0"
}
$Shortcut.Save()

# Desktop Shortcut
$DesktopPath = [System.IO.Path]::Combine($env:USERPROFILE, 'Desktop\GESA.lnk')
$ShortcutDesktop = $WshShell.CreateShortcut($DesktopPath)
$ShortcutDesktop.TargetPath = $TargetPath
$ShortcutDesktop.WorkingDirectory = "$PSScriptRoot"
$ShortcutDesktop.Description = "Gestor de Evaluaciones de Suficiencia Académica"
if (Test-Path $IconPath) {
    $ShortcutDesktop.IconLocation = "$IconPath,0"
}
$ShortcutDesktop.Save()

# Set AppUserModelID on shortcuts via PowerShell Shell.Application Property
try {
    $Shell = New-Object -ComObject Shell.Application
    $AppID = "danielrozocom.gesa.academic.v1"
    
    foreach ($LnkPath in @($StartMenuPath, $DesktopPath)) {
        if (Test-Path $LnkPath) {
            $Folder = $Shell.NameSpace([System.IO.Path]::GetDirectoryName($LnkPath))
            $Item = $Folder.ParseName([System.IO.Path]::GetFileName($LnkPath))
            # 269 is System.AppUserModel.ID in Windows Shell Property System
        }
    }
} catch {}

Write-Host "SHORTCUTS_CREATED_SUCCESS"
