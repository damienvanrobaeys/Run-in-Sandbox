$Sandbox_Folder = "C:\Users\WDAGUtilityAccount\Desktop\Run_in_Sandbox"

$FileName = (Get-Item $ScriptPath).BaseName

New-Item "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs" -Force -Type Directory

Set-Location "$Intunewin_Extracted_Folder\$FileName"
$file = "$Sandbox_Folder\EXE_Command_File.txt"
& { Invoke-Expression (Get-Content -Raw $file) }