#***************************************************************************************************************
# Author: Damien VAN ROBAEYS
# Website: http://www.systanddeploy.com
# Twitter: https://twitter.com/syst_and_deploy
# Purpose: This script will remove context menus added to run quickly files in Windows Sandbox
#***************************************************************************************************************

$Current_Folder = split-path $MyInvocation.MyCommand.Path
$ProgData = $env:ProgramData
$Sandbox_Folder = "$ProgData\Run_in_Sandbox"
Remove-item $Sandbox_Folder -recurse -force

$List_Drive = get-psdrive | where {$_.Name -eq "HKCR_SD"}
If($List_Drive -ne $null){Remove-PSDrive $List_Drive}
New-PSDrive -PSProvider registry -Root HKEY_CLASSES_ROOT -Name HKCR_SD

# REMOVE RUN ON PS1
$PS1_Shell_Registry_Key = "HKCR_SD:\Microsoft.PowerShellScript.1\Shell"
$PS1_Basic_Run = "Run the PS1 in Sandbox"
$PS1_Parameter_Run = "Run the PS1 in Sandbox with parameters"
Remove-Item -Path "$PS1_Shell_Registry_Key\$PS1_Basic_Run" -Recurse
Remove-Item -Path "$PS1_Shell_Registry_Key\$PS1_Parameter_Run" -Recurse

# REMOVE RUN ON VBS
$VBS_Shell_Registry_Key = "HKCR_SD:\VBSFile\Shell"
$VBS_Basic_Run = "Run the VBS in Sandbox"
$VBS_Parameter_Run = "Run the VBS in Sandbox with parameters"
Remove-Item -Path "$VBS_Shell_Registry_Key\$VBS_Basic_Run" -Recurse
Remove-Item -Path "$VBS_Shell_Registry_Key\$VBS_Parameter_Run" -Recurse

# REMOVE RUN ON EXE
$EXE_Shell_Registry_Key = "HKCR_SD:\exefile\Shell"
$EXE_Basic_Run = "Run the EXE in Sandbox"
Remove-Item -Path "$EXE_Shell_Registry_Key\$EXE_Basic_Run" -Recurse

# RUN ON MSI
$MSI_Shell_Registry_Key = "HKCR_SD:\Msi.Package\Shell"
$MSI_Basic_Run = "Run the MSI in Sandbox"
Remove-Item -Path "$MSI_Shell_Registry_Key\$MSI_Basic_Run" -Recurse

# RUN ON ZIP
$ZIP_Shell_Registry_Key = "HKCR_SD:\CompressedFolder\Shell"
$ZIP_Basic_Run = "Extract the ZIP in Sandbox"
Remove-Item -Path "$ZIP_Shell_Registry_Key\$ZIP_Basic_Run" -Recurse

# RUN ON ZIP if WinRAR is installed
If(test-path "HKCR_SD:\WinRAR.ZIP\Shell")
	{
		$ZIP_WinRAR_Shell_Registry_Key = "HKCR_SD:\WinRAR.ZIP\Shell"
		$ZIP_WinRAR_Basic_Run = "Extract the ZIP in Sandbox"
		Remove-Item -Path "$ZIP_WinRAR_Shell_Registry_Key\$ZIP_WinRAR_Basic_Run" -Recurse
	}

If($List_Drive -ne $null){Remove-PSDrive $List_Drive}
