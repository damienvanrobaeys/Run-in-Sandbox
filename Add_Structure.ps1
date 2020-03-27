#***************************************************************************************************************
# Author: Damien VAN ROBAEYS
# Website: http://www.systanddeploy.com
# Twitter: https://twitter.com/syst_and_deploy
#***************************************************************************************************************

$Current_Folder = split-path $MyInvocation.MyCommand.Path
$Sources = $Current_Folder + "\" + "Sources\*"

$ProgData = $env:ProgramData
$Destination_folder = $ProgData
copy-item $Sources $Destination_folder -force -recurse


# RUN ON PS1
$PS1_Shell_Registry_Key = "HKCR_SD:\Microsoft.PowerShellScript.1\Shell"
$PS1_Basic_Run = "Run the PS1 in Sandbox"
$PS1_Parameter_Run = "Run the PS1 in Sandbox with parameters"
$ContextMenu_Basic_PS1 = "$PS1_Shell_Registry_Key\$PS1_Basic_Run"
$ContextMenu_Parameters_PS1 = "$PS1_Shell_Registry_Key\$PS1_Parameter_Run"
$List_Drive = get-psdrive | where {$_.Name -eq "HKCR_SD"}
If($List_Drive -ne $null){Remove-PSDrive $List_Drive}
New-PSDrive -PSProvider registry -Root HKEY_CLASSES_ROOT -Name HKCR_SD
New-Item -Path $PS1_Shell_Registry_Key -Name $PS1_Basic_Run -force
New-Item -Path $PS1_Shell_Registry_Key -Name $PS1_Parameter_Run -force
New-Item -Path $ContextMenu_Basic_PS1 -Name "Command" -force
New-Item -Path $ContextMenu_Parameters_PS1 -Name "Command" -force
$Command_For_Basic_PS1 = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type PS1Basic -LiteralPath "%V" -ScriptPath "%V"' 
$Command_For_Params_PS1 = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type PS1Params -LiteralPath "%V" -ScriptPath "%V"' 
Set-Item -Path "$ContextMenu_Basic_PS1\command" -Value $Command_For_Basic_PS1 -force
Set-Item -Path "$ContextMenu_Parameters_PS1\command" -Value $Command_For_Params_PS1 -force


# RUN ON VBS
$VBS_Shell_Registry_Key = "HKCR_SD:\VBSFile\Shell"
$VBS_Basic_Run = "Run the VBS in Sandbox"
$VBS_Parameter_Run = "Run the VBS in Sandbox with parameters"
$ContextMenu_Basic_VBS = "$VBS_Shell_Registry_Key\$VBS_Basic_Run"
$ContextMenu_Parameters_VBS = "$VBS_Shell_Registry_Key\$VBS_Parameter_Run"
New-Item -Path $VBS_Shell_Registry_Key -Name $VBS_Basic_Run -force
New-Item -Path $VBS_Shell_Registry_Key -Name $VBS_Parameter_Run -force
New-Item -Path $ContextMenu_Basic_VBS -Name "Command" -force
New-Item -Path $ContextMenu_Parameters_VBS -Name "Command" -force
$Command_For_Basic_VBS = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type VBSBasic -LiteralPath "%V" -ScriptPath "%V"' 
$Command_For_Params_VBS = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type VBSParams -LiteralPath "%V" -ScriptPath "%V"' 
Set-Item -Path "$ContextMenu_Basic_VBS\command" -Value $Command_For_Basic_VBS -force
Set-Item -Path "$ContextMenu_Parameters_VBS\command" -Value $Command_For_Params_VBS -force


# RUN ON EXE
$EXE_Shell_Registry_Key = "HKCR_SD:\exefile\Shell"
$EXE_Basic_Run = "Run the EXE in Sandbox"
$ContextMenu_Basic_EXE = "$EXE_Shell_Registry_Key\$EXE_Basic_Run"
New-Item -Path $EXE_Shell_Registry_Key -Name $EXE_Basic_Run -force
New-Item -Path $ContextMenu_Basic_EXE -Name "Command" -force
$Command_For_EXE = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type EXE -LiteralPath "%V" -ScriptPath "%V"' 
Set-Item -Path "$ContextMenu_Basic_EXE\command" -Value $Command_For_EXE -force


# RUN ON MSI
$MSI_Shell_Registry_Key = "HKCR_SD:\Msi.Package\Shell"
$MSI_Basic_Run = "Run the MSI in Sandbox"
$ContextMenu_Basic_MSI = "$MSI_Shell_Registry_Key\$MSI_Basic_Run"
New-Item -Path $MSI_Shell_Registry_Key -Name $MSI_Basic_Run -force
New-Item -Path $ContextMenu_Basic_MSI -Name "Command" -force
$Command_For_MSI = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type MSI -LiteralPath "%V" -ScriptPath "%V"' 
Set-Item -Path "$ContextMenu_Basic_MSI\command" -Value $Command_For_MSI -force











