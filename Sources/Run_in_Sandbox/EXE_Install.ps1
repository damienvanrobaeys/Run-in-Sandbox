$Sandbox_Folder = "C:\Run_in_Sandbox"
$FolderPath = Split-Path (Split-Path "$ScriptPath" -Parent) -Leaf
$DirectoryName = (get-item $ScriptPath).DirectoryName
$FileName = (get-item $ScriptPath).BaseName

New-item "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs" -Force -Type Directory

Set-Location "$Intunewin_Extracted_Folder\$FileName"
$File = "$Sandbox_Folder\EXE_Command_File.txt"
& { Invoke-Expression (Get-Content -Raw $File)}