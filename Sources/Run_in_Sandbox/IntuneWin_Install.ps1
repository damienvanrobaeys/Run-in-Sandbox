Param (
	[String]$Intunewin_Content_File = "C:\Run_in_Sandbox\Intunewin_Folder.txt",
	[String]$Intunewin_Command_File = "C:\Run_in_Sandbox\Intunewin_Install_Command.txt"
)
if (-not (Test-Path $Intunewin_Content_File) ){
	EXIT
}
if (-not (Test-Path $Intunewin_Command_File) ){
	EXIT
}

$Sandbox_Folder = "C:\Run_in_Sandbox"
$ScriptPath = Get-Content -Raw $Intunewin_Content_File
$Command = Get-Content -Raw $Intunewin_Command_File
$Command = $Command.replace('"','')

$FileName = (Get-Item $ScriptPath).BaseName

$Intunewin_Extracted_Folder = "C:\Windows\Temp\intunewin"
New-Item $Intunewin_Extracted_Folder -Type Directory -Force | Out-Null
Copy-Item $ScriptPath $Intunewin_Extracted_Folder -Force
$New_Intunewin_Path = "$Intunewin_Extracted_Folder\$FileName.intunewin"

Set-Location $Sandbox_Folder
& .\IntuneWinAppUtilDecoder.exe $New_Intunewin_Path -s
$IntuneWinDecoded_File_Name = "$Intunewin_Extracted_Folder\$FileName.Intunewin.decoded"

New-Item "$Intunewin_Extracted_Folder\$FileName" -Type Directory -Force | Out-Null

$IntuneWin_Rename = "$FileName.zip"

Rename-Item $IntuneWinDecoded_File_Name $IntuneWin_Rename -Force

$Extract_Path = "$Intunewin_Extracted_Folder\$FileName"
Expand-Archive -LiteralPath "$Intunewin_Extracted_Folder\$IntuneWin_Rename" -DestinationPath $Extract_Path -Force

Remove-Item "$Intunewin_Extracted_Folder\$IntuneWin_Rename" -Force
Start-Sleep 1

$ServiceUI = "$Sandbox_Folder\ServiceUI.exe"
$PsExec = "$Sandbox_Folder\PSTools\PsExec64.exe"

$cmd = "$PsExec \\localhost -w $Extract_Path -nobanner -accepteula -i -s C:\Windows\System32\WindowsPowershell\v1.0\powershell.exe -ExecutionPolicy Unrestricted -NoProfile -Command '$Command'"
$cmd = "Write-Host `"Installing....`"; $cmd"

& $ServiceUI -process:explorer.exe C:\Windows\System32\WindowsPowershell\v1.0\powershell.exe -Command $cmd