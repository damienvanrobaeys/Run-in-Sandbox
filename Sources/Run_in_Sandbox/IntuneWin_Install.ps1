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

$Intunewin_Content_File = "$Sandbox_Folder\Intunewin_Folder.txt"
$Intunewin_FilePath = get-content $Intunewin_Content_File
$Intunewin_FileName = $Intunewin_FilePath.split("\")[-1]

copy-item $Intunewin_FilePath $Intunewin_Extracted_Folder -Force
$Intunewin_New_path = "$Intunewin_Extracted_Folder\$Intunewin_FileName"

& .\IntuneWinAppUtilDecoder.exe $Intunewin_New_path -s	

$Intunewin_Extract_Directory = (get-item $Intunewin_New_path).Directory	
$IntuneWin_File_Name = (get-item $Intunewin_New_path).BaseName		
$IntuneWinDecoded_File = "$Intunewin_Extract_Directory\$IntuneWin_File_Name.decoded.zip"	

$Extract_Folder = (get-item $Intunewin_New_path).BaseName
$Extract_DirectoryName = (get-item $Intunewin_New_path).DirectoryName

Expand-Archive -LiteralPath $IntuneWinDecoded_File -DestinationPath "$Extract_DirectoryName\$Extract_Folder" -Force

$ServiceUI = "$Sandbox_Folder\ServiceUI.exe"
$PsExec = "$Sandbox_Folder\PSTools\PsExec64.exe"

$cmd = "$PsExec \\localhost -w $Extract_Path -nobanner -accepteula -i -s C:\Windows\System32\WindowsPowershell\v1.0\powershell.exe -ExecutionPolicy Unrestricted -NoProfile -Command '$Command'"
$cmd = "Write-Host `"Installing....`"; $cmd"

& $ServiceUI -process:explorer.exe C:\Windows\System32\WindowsPowershell\v1.0\powershell.exe -Command $cmd