#***************************************************************************************************************
# Author: Damien VAN ROBAEYS
# Website: http://www.systanddeploy.com
# Twitter: https://twitter.com/syst_and_deploy
#***************************************************************************************************************
$Progress_Activity = "Enabling Run in Sandbox context menus"
write-progress -activity $Progress_Activity -percentcomplete 1;

$Current_Folder = split-path $MyInvocation.MyCommand.Path
$Sources = $Current_Folder + "\" + "Sources\*"

$ProgData = $env:ProgramData
$Destination_folder = $ProgData
copy-item $Sources $Destination_folder -force -recurse
$Sandbox_Icon = "$Destination_folder\Run_in_Sandbox\sandbox.ico"

$Run_in_Sandbox_Folder = "$ProgData\Sources\Run_in_Sandbox\Run_in_Sandbox"
$XML_Config = "$Current_Folder\Sources\Run_in_Sandbox\Sandbox_Config.xml"
$Get_XML_Content = [xml] (Get-Content $XML_Config)

$List_Drive = get-psdrive | where {$_.Name -eq "HKCR_SD"}
If($List_Drive -ne $null){Remove-PSDrive $List_Drive}
New-PSDrive -PSProvider registry -Root HKEY_CLASSES_ROOT -Name HKCR_SD | out-null

write-progress -activity $Progress_Activity  -percentcomplete 10;

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


write-progress -activity $Progress_Activity -percentcomplete 20;


# RUN ON PS1
$PS1_Shell_Registry_Key = "HKCR_SD:\Microsoft.PowerShellScript.1\Shell"
$PS1_Basic_Run = $Get_Language_File_Content.PowerShell.Basic
$PS1_Parameter_Run = $Get_Language_File_Content.PowerShell.Parameters
$ContextMenu_Basic_PS1 = "$PS1_Shell_Registry_Key\$PS1_Basic_Run"
$ContextMenu_Parameters_PS1 = "$PS1_Shell_Registry_Key\$PS1_Parameter_Run"

New-Item -Path $PS1_Shell_Registry_Key -Name $PS1_Basic_Run -force | out-null
New-Item -Path $PS1_Shell_Registry_Key -Name $PS1_Parameter_Run -force | out-null
New-Item -Path $ContextMenu_Basic_PS1 -Name "Command" -force | out-null
New-Item -Path $ContextMenu_Parameters_PS1 -Name "Command" -force | out-null
$Command_For_Basic_PS1 = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type PS1Basic -LiteralPath "%V" -ScriptPath "%V"' 
$Command_For_Params_PS1 = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type PS1Params -LiteralPath "%V" -ScriptPath "%V"' 
# Set the command path
Set-Item -Path "$ContextMenu_Basic_PS1\command" -Value $Command_For_Basic_PS1 -force | out-null
Set-Item -Path "$ContextMenu_Parameters_PS1\command" -Value $Command_For_Params_PS1 -force | out-null
# Add Sandbox Icons
New-ItemProperty -Path $ContextMenu_Basic_PS1 -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null
New-ItemProperty -Path $ContextMenu_Parameters_PS1 -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null


write-progress -activity $Progress_Activity  -percentcomplete 30;


# RUN ON VBS
$VBS_Shell_Registry_Key = "HKCR_SD:\VBSFile\Shell"
$VBS_Basic_Run = $Get_Language_File_Content.VBS.Basic
$VBS_Parameter_Run = $Get_Language_File_Content.VBS.Parameters
$ContextMenu_Basic_VBS = "$VBS_Shell_Registry_Key\$VBS_Basic_Run"
$ContextMenu_Parameters_VBS = "$VBS_Shell_Registry_Key\$VBS_Parameter_Run"

New-Item -Path $VBS_Shell_Registry_Key -Name $VBS_Basic_Run -force | out-null
New-Item -Path $VBS_Shell_Registry_Key -Name $VBS_Parameter_Run -force | out-null
New-Item -Path $ContextMenu_Basic_VBS -Name "Command" -force | out-null
New-Item -Path $ContextMenu_Parameters_VBS -Name "Command" -force | out-null
$Command_For_Basic_VBS = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type VBSBasic -LiteralPath "%V" -ScriptPath "%V"' 
$Command_For_Params_VBS = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type VBSParams -LiteralPath "%V" -ScriptPath "%V"' 
# Set the command path
Set-Item -Path "$ContextMenu_Basic_VBS\command" -Value $Command_For_Basic_VBS -force | out-null
Set-Item -Path "$ContextMenu_Parameters_VBS\command" -Value $Command_For_Params_VBS -force | out-null
# Add Sandbox Icons
New-ItemProperty -Path $ContextMenu_Basic_VBS -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null
New-ItemProperty -Path $ContextMenu_Parameters_VBS -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null


write-progress -activity $Progress_Activity  -percentcomplete 40;


# RUN ON EXE
$EXE_Shell_Registry_Key = "HKCR_SD:\exefile\Shell"
$EXE_Basic_Run = $Get_Language_File_Content.EXE
$ContextMenu_Basic_EXE = "$EXE_Shell_Registry_Key\$EXE_Basic_Run"

New-Item -Path $EXE_Shell_Registry_Key -Name $EXE_Basic_Run -force | out-null
New-Item -Path $ContextMenu_Basic_EXE -Name "Command" -force | out-null
# Add Sandbox Icons
New-ItemProperty -Path $ContextMenu_Basic_EXE -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null
$Command_For_EXE = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type EXE -LiteralPath "%V" -ScriptPath "%V"' 
# Set the command path
Set-Item -Path "$ContextMenu_Basic_EXE\command" -Value $Command_For_EXE -force | out-null


write-progress -activity $Progress_Activity  -percentcomplete 50;


# RUN ON MSI
$MSI_Shell_Registry_Key = "HKCR_SD:\Msi.Package\Shell"
$MSI_Basic_Run = $Get_Language_File_Content.MSI
$ContextMenu_Basic_MSI = "$MSI_Shell_Registry_Key\$MSI_Basic_Run"

New-Item -Path $MSI_Shell_Registry_Key -Name $MSI_Basic_Run -force | out-null
New-Item -Path $ContextMenu_Basic_MSI -Name "Command" -force | out-null
# Add Sandbox Icons
New-ItemProperty -Path $ContextMenu_Basic_MSI -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null
$Command_For_MSI = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type MSI -LiteralPath "%V" -ScriptPath "%V"' 
# Set the command path
Set-Item -Path "$ContextMenu_Basic_MSI\command" -Value $Command_For_MSI -force | out-null


write-progress -activity $Progress_Activity  -percentcomplete 60;


# RUN ON ZIP
$ZIP_Shell_Registry_Key = "HKCR_SD:\CompressedFolder\Shell"
$ZIP_Basic_Run = $Get_Language_File_Content.ZIP
$ContextMenu_Basic_ZIP = "$ZIP_Shell_Registry_Key\$ZiP_Basic_Run"

New-Item -Path $ZIP_Shell_Registry_Key -Name $ZiP_Basic_Run -force | out-null
New-Item -Path $ContextMenu_Basic_ZIP -Name "Command" -force | out-null
# Add Sandbox Icons
New-ItemProperty -Path $ContextMenu_Basic_ZIP -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null
$Command_For_ZIP = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type ZIP -LiteralPath "%V" -ScriptPath "%V"' 
# Set the command path
Set-Item -Path "$ContextMenu_Basic_ZIP\command" -Value $Command_For_ZIP -force | out-null


write-progress -activity $Progress_Activity  -percentcomplete 70;


# RUN ON ZIP if WinRAR is installed
If(test-path "HKCR_SD:\WinRAR.ZIP\Shell")
	{
		$ZIP_WinRAR_Shell_Registry_Key = "HKCR_SD:\WinRAR.ZIP\Shell"
		$ZIP_WinRAR_Basic_Run = $Get_Language_File_Content.ZIP				
		$ContextMenu_Basic_ZIP_RAR = "$ZIP_WinRAR_Shell_Registry_Key\$ZIP_WinRAR_Basic_Run"
		
		New-Item -Path $ZIP_WinRAR_Shell_Registry_Key -Name $ZIP_WinRAR_Basic_Run -force | out-null
		New-Item -Path $ContextMenu_Basic_ZIP_RAR -Name "Command" -force | out-null
		# Add Sandbox Icons
		New-ItemProperty -Path $ContextMenu_Basic_ZIP_RAR -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null
		$Command_For_ZIP = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type ZIP -LiteralPath "%V" -ScriptPath "%V"' 
		# Set the command path
		Set-Item -Path "$ContextMenu_Basic_ZIP_RAR\command" -Value $Command_For_ZIP -force | out-null
	}
	
	
write-progress -activity $Progress_Activity  -percentcomplete 80;	


# Share this folder - Inside the folder
$Folder_Inside_Shell_Registry_Key = "HKCR_SD:\Directory\Background\shell"
$Folder_Inside_Basic_Run = $Get_Language_File_Content.Folder				

$ContextMenu_Folder_Inside = "$Folder_Inside_Shell_Registry_Key\$Folder_Inside_Basic_Run"

New-Item -Path $Folder_Inside_Shell_Registry_Key -Name $Folder_Inside_Basic_Run -force | out-null
New-Item -Path $ContextMenu_Folder_Inside -Name "Command" -force | out-null
# Add Sandbox Icons
New-ItemProperty -Path $ContextMenu_Folder_Inside -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null
$Command_For_Folder_Inside = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type Folder_Inside -LiteralPath "%V" -ScriptPath "%V"' 
# Set the command path
Set-Item -Path "$ContextMenu_Folder_Inside\command" -Value $Command_For_Folder_Inside -force | out-null


write-progress -activity $Progress_Activity  -percentcomplete 90;	


# Share this folder - Right-click on the folder
$Folder_On_Shell_Registry_Key = "HKCR_SD:\Directory\shell"
$Folder_On_Run = $Get_Language_File_Content.Folder				

$ContextMenu_Folder_On = "$Folder_On_Shell_Registry_Key\$Folder_On_Run"

New-Item -Path $Folder_On_Shell_Registry_Key -Name $Folder_On_Run -force | out-null
New-Item -Path $ContextMenu_Folder_On -Name "Command" -force | out-null
# Add Sandbox Icons
New-ItemProperty -Path $ContextMenu_Folder_On -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null
$Command_For_Folder_On = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type Folder_On -LiteralPath "%V" -ScriptPath "%V"' 
# Set the command path
Set-Item -Path "$ContextMenu_Folder_On\command" -Value $Command_For_Folder_On -force | out-null


If($List_Drive -ne $null){Remove-PSDrive $List_Drive}

write-progress -activity $Progress_Activity  -percentcomplete 100;	