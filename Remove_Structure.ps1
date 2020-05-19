#***************************************************************************************************************
# Author: Damien VAN ROBAEYS
# Website: http://www.systanddeploy.com
# Twitter: https://twitter.com/syst_and_deploy
# Purpose: This script will remove context menus added to run quickly files in Windows Sandbox
#***************************************************************************************************************

$Current_Folder = split-path $MyInvocation.MyCommand.Path
$ProgData = $env:ProgramData
$Sandbox_Folder = "$ProgData\Run_in_Sandbox"

$XML_Config = "$Sandbox_Folder\Sandbox_Config.xml"
$Get_XML_Content = [xml] (Get-Content $XML_Config)
$Main_Language = $Get_XML_Content.Configuration.Main_Language
If($Main_Language -ne $null)
	{
		$Get_lang_to_install = $Get_XML_Content.Configuration.Main_Language	
	}
Else
	{
		$Get_lang_to_install = (Get-Culture).name
	}
$Language_File = (Get-Childitem "$Current_Folder\Sources\Run_in_Sandbox\Languages_XML" | Where {$_.Basename -like "*$Get_lang_to_install*"}).fullname
$Get_Language_File_Content = ([xml](get-content $Language_File)).Configuration


$List_Drive = get-psdrive | where {$_.Name -eq "HKCR_SD"}
If($List_Drive -ne $null){Remove-PSDrive $List_Drive}
New-PSDrive -PSProvider registry -Root HKEY_CLASSES_ROOT -Name HKCR_SD


# REMOVE RUN ON PS1
$PS1_Shell_Registry_Key = "HKCR_SD:\Microsoft.PowerShellScript.1\Shell"
$PS1_Basic_Run = $Get_Language_File_Content.PowerShell.Basic
$PS1_Parameter_Run = $Get_Language_File_Content.PowerShell.Parameters
Remove-Item -Path "$PS1_Shell_Registry_Key\$PS1_Basic_Run" -Recurse
Remove-Item -Path "$PS1_Shell_Registry_Key\$PS1_Parameter_Run" -Recurse


# REMOVE RUN ON VBS
$VBS_Shell_Registry_Key = "HKCR_SD:\VBSFile\Shell"
$VBS_Basic_Run = $Get_Language_File_Content.VBS.Basic
$VBS_Parameter_Run = $Get_Language_File_Content.VBS.Parameters
Remove-Item -Path "$VBS_Shell_Registry_Key\$VBS_Basic_Run" -Recurse
Remove-Item -Path "$VBS_Shell_Registry_Key\$VBS_Parameter_Run" -Recurse


# REMOVE RUN ON EXE
$EXE_Shell_Registry_Key = "HKCR_SD:\exefile\Shell"
$EXE_Basic_Run = $Get_Language_File_Content.EXE
Remove-Item -Path "$EXE_Shell_Registry_Key\$EXE_Basic_Run" -Recurse


# RUN ON MSI
$MSI_Shell_Registry_Key = "HKCR_SD:\Msi.Package\Shell"
$MSI_Basic_Run = $Get_Language_File_Content.MSI
Remove-Item -Path "$MSI_Shell_Registry_Key\$MSI_Basic_Run" -Recurse


# RUN ON ZIP
$ZIP_Shell_Registry_Key = "HKCR_SD:\CompressedFolder\Shell"
$ZIP_Basic_Run = $Get_Language_File_Content.ZIP
Remove-Item -Path "$ZIP_Shell_Registry_Key\$ZIP_Basic_Run" -Recurse


# RUN ON ZIP if WinRAR is installed
If(test-path "HKCR_SD:\WinRAR.ZIP\Shell")
	{
		$ZIP_WinRAR_Shell_Registry_Key = "HKCR_SD:\WinRAR.ZIP\Shell"
		$ZIP_WinRAR_Basic_Run = $Get_Language_File_Content.ZIP	
		Remove-Item -Path "$ZIP_WinRAR_Shell_Registry_Key\$ZIP_WinRAR_Basic_Run" -Recurse
	}
	
	
# Share this folder - Inside the folder
$Folder_Inside_Shell_Registry_Key = "HKCR_SD:\Directory\Background\shell"
$Folder_Inside_Basic_Run = $Get_Language_File_Content.Folder				
Remove-Item -Path "$Folder_Inside_Shell_Registry_Key\$Folder_Inside_Basic_Run" -Recurse


# Share this folder - Right-click on the folder
$Folder_On_Shell_Registry_Key = "HKCR_SD:\Directory\shell"
$Folder_On_Run = $Get_Language_File_Content.Folder				
Remove-Item -Path "$Folder_On_Shell_Registry_Key\$Folder_On_Run" -Recurse	

If($List_Drive -ne $null){Remove-PSDrive $List_Drive}

Remove-item $Sandbox_Folder -recurse -force