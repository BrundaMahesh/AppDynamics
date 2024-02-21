'Create a FileSystemObject
Set fsObject = CreateObject("Scripting.FileSystemObject")

'Get the directory of the current script
scriptDir = fsObject.GetParentFolderName(WScript.ScriptFullName)

'Execute VBScript file
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cscript """ & scriptDir & "\InstallService.vbs""", 1, True
Set WshShell = Nothing

'Execute PowerShell script
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "powershell.exe -ExecutionPolicy RemoteSigned -Command ""&{cd '" & scriptDir & "'; & '.\Windows_Path_Enumerate.ps1' -FixUninstall -FixEnv}""", 1, True
Set WshShell = Nothing



