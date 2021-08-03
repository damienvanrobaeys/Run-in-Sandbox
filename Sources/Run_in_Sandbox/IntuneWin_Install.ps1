$Sandbox_Folder = "C:\Users\WDAGUtilityAccount\Desktop\Run_in_Sandbox"
$Intunewin_Content_File = "$Sandbox_Folder\Intunewin_Folder.txt"
$ScriptPath = get-content $Intunewin_Content_File

$FolderPath = Split-Path (Split-Path "$ScriptPath" -Parent) -Leaf
$DirectoryName = (get-item $ScriptPath).DirectoryName
$FileName = (get-item $ScriptPath).BaseName

$Intunewin_Extracted_Folder = "C:\Windows\Temp\intunewin"
new-item $Intunewin_Extracted_Folder -Type Directory -Force
copy-item $ScriptPath $Intunewin_Extracted_Folder -Force
$New_Intunewin_Path = "$Intunewin_Extracted_Folder\$FileName.intunewin"

set-location $Sandbox_Folder
& .\IntuneWinAppUtilDecoder.exe $New_Intunewin_Path -s	
$IntuneWinDecoded_File_Name = "$Intunewin_Extracted_Folder\$FileName.Intunewin.decoded"	
	
new-item "$Intunewin_Extracted_Folder\$FileName" -Type Directory -Force | out-null

$IntuneWin_Rename = "$FileName.zip"

Rename-Item $IntuneWinDecoded_File_Name $IntuneWin_Rename -force

$Extract_Path = "$Intunewin_Extracted_Folder\$FileName"
Expand-Archive -LiteralPath "$Intunewin_Extracted_Folder\$IntuneWin_Rename" -DestinationPath $Extract_Path -Force

Remove-Item "$Intunewin_Extracted_Folder\$IntuneWin_Rename" -force
sleep 1

set-location "$Intunewin_Extracted_Folder\$FileName"
$file = "$Sandbox_Folder\Intunewin_Install_Command.txt"
& { Invoke-Expression (Get-Content -Raw $file) }

