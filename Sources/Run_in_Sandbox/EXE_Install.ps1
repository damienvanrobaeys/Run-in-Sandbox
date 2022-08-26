$Sandbox_Folder = "C:\Users\WDAGUtilityAccount\Desktop\Run_in_Sandbox"

$FolderPath = Split-Path (Split-Path "$ScriptPath" -Parent) -Leaf
$DirectoryName = (get-item $ScriptPath).DirectoryName
$FileName = (get-item $ScriptPath).BaseName

New-item "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs" -Force -Type Directory

set-location "$Intunewin_Extracted_Folder\$FileName"
$file = "$Sandbox_Folder\EXE_Command_File.txt"
& { Invoke-Expression (Get-Content -Raw $file) }