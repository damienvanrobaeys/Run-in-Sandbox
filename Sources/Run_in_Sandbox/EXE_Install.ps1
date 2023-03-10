$Sandbox_Folder = "C:\Run_in_Sandbox"
$Exe_Content_File = "$Sandbox_Folder\EXE_Command_File.txt"
$ExeCommand = Get-Content $Exe_Content_File
$Full_Exe_Path = $ExeCommand -Split "(?<=.exe)\s" | Select-Object -First 1
$Directory_Path = (Get-Item $Full_Exe_Path).DirectoryName

Set-Location "$Directory_Path"
& { Invoke-Expression $ExeCommand }