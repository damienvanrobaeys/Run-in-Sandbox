$Sandbox_Folder = "C:\Run_in_Sandbox"
$File = "$Sandbox_Folder\EXE_Command_File.txt"

$Content = Get-Content -Raw $File
$BaseFolder = Split-Path($Content.Split('"')[1])
$Executable = Split-Path($Content.Split('"')[1]) -Leaf
$Executable = "./$Executable"
$Arguments = ($Content.Split('"',3)[-1]).Trim()

# The executable has to be quoted, otherwise Invoke-Expression breaks on file names with spaces
$Command = "& `"$Executable`" " + $Arguments

Set-Location -Path $BaseFolder

& { Invoke-Expression $Command}