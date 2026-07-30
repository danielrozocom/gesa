Set fso = CreateObject("Scripting.FileSystemObject")
currentDir = fso.GetParentFolderName(WScript.ScriptFullName)

Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = currentDir
WshShell.Run chr(34) & currentDir & "\start.bat" & Chr(34), 0

Set WshShell = Nothing
