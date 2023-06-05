$Intunewin_Extracted_Folder = "C:\Windows\Temp\intunewin"
$Sandbox_Folder = "C:\Users\WDAGUtilityAccount\Desktop\Run_in_Sandbox"

New-Item $Intunewin_Extracted_Folder -Type Directory -Force
New-item "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs" -Force -Type Directory
set-location $Sandbox_Folder

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

Remove-Item $IntuneWinDecoded_File

set-location "$Intunewin_Extracted_Folder\$Extract_Folder"
$file = "$Sandbox_Folder\Intunewin_Install_Command.txt"
& { Invoke-Expression (Get-Content -Raw $file)}