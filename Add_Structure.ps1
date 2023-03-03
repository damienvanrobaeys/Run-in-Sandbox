Param (
	[Switch]$NoSilent
)

#***************************************************************************************************************
# Author: Damien VAN ROBAEYS
# Website: http://www.systanddeploy.com
# Twitter: https://twitter.com/syst_and_deploy
#***************************************************************************************************************
$TEMP_Folder = $env:temp
$Log_File = "$TEMP_Folder\RunInSandbox_Install.log"
$Current_Folder = Split-Path $MyInvocation.MyCommand.Path

If (Test-Path $Log_File) { Remove-Item $Log_File }
New-Item $Log_File -type file -Force | Out-Null
Function Write_Log {
	param (
		$Message_Type,
		$Message
	)

	$MyDate = "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
	Add-Content $Log_File "$MyDate - $Message_Type : $Message"
	Write-Output "$MyDate - $Message_Type : $Message"
}

Function Export_Reg_Config {
	param (
		$Reg_Path,
		$Backup_Path
	)

	Write_Log -Message_Type "INFO" -Message "Exporting registry path: $Reg_Path"

	If (Test-Path "HKCR_SD:\$Reg_Path") {
		Try {
			reg export "HKEY_CLASSES_ROOT\$Reg_Path" $Backup_Path /y | Out-Null
			Write_Log -Message_Type "SUCCESS" -Message "$Reg_Path has been exported"
		}
		Catch {
			Write_Log -Message_Type "ERROR" -Message "$Reg_Path has not been exported"
		}
	}
	Else {
		Write_Log -Message_Type "INFO" -Message "Can not find registry path: HKCR_SD:\$Reg_Path"
	}
	Add-Content $log_file ""
}


Write_Log -Message_Type "INFO" -Message "Starting the configuration of RunInSandbox"

[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
$Run_As_Admin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
If ($Run_As_Admin -eq $False) {
	Write_Log -Message_Type "ERROR" -Message "The script has not been lauched with admin rights"
	[System.Windows.Forms.MessageBox]::Show("Please run the tool with admin rights :-)")
	break
}
Write_Log -Message_Type "INFO" -Message "The script has been launched with admin rights"

$Is_Sandbox_Installed = (Get-WindowsOptionalFeature -Online | Where-Object { $_.featurename -eq "Containers-DisposableClientVM" }).state
If ($Is_Sandbox_Installed -eq "Disabled") {
	Write_Log -Message_Type "ERROR" -Message "The feature Windows Sandbox is not installed !!!"
	[System.Windows.Forms.MessageBox]::Show("The feature Windows Sandbox is not installed !!!")
	break
}

Write_Log -Message_Type "INFO" -Message "The Windows Sandbox feature is installed"

$Current_Folder = Split-Path $MyInvocation.MyCommand.Path
$Sources = $Current_Folder + "\" + "Sources\*"
If (-not (Test-Path $Sources)) {
	Write_Log -Message_Type "ERROR" -Message "Sources folder is missing"
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
	[System.Windows.Forms.MessageBox]::Show("It seems you don't have dowloaded all the folder structure.`nThe folder Sources is missing !!!")
	break
}

Write_Log -Message_Type "INFO" -Message "The sources folder exists"

Add-Content $log_file ""

$Progress_Activity = "Enabling Run in Sandbox context menus"
Write-Progress -Activity $Progress_Activity -PercentComplete 1

$Check_Sources_Files_Count = (Get-ChildItem "$Current_Folder\Sources\Run_in_Sandbox" -Recurse).count
If ($Check_Sources_Files_Count -ne 58) {
	Write_Log -Message_Type "ERROR" -Message "Some contents are missing"
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
	[System.Windows.Forms.MessageBox]::Show("It seems you don't have dowloaded all the folder structure !!!")
	break
}


$ProgData = $env:ProgramData
$Destination_folder = "$ProgData\Run_in_Sandbox"
Try {
	Copy-Item $Sources $ProgData -Force -Recurse | Out-Null
	Write_Log -Message_Type "SUCCESS" -Message "Sources have been copied in $ProgData\Run_in_Sandbox"
}
Catch {
	Write_Log -Message_Type "ERROR" -Message "Sources have not been copied in $ProgData\Run_in_Sandbox"
	Break
}

$Sources_Unblocked = $False
Try {
	Get-ChildItem -Recurse $Destination_folder | Unblock-File
	Write_Log -Message_Type "SUCCESS" -Message "Sources files have been unblocked"
	$Sources_Unblocked = $True
}
Catch {
	Write_Log -Message_Type "ERROR" -Message "Sources files have not been unblocked"
	Break
}

If ($Sources_Unblocked -ne $True) {
	Write_Log -Message_Type "ERROR" -Message "Source files could not be unblocked"
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
	[System.Windows.Forms.MessageBox]::Show("Source files could not be unblocked")
	break
}

$Script:Current_User_SID = (Get-ChildItem Registry::\HKEY_USERS | Where-Object { Test-Path "$($_.pspath)\Volatile Environment" } | ForEach-Object { (Get-ItemProperty "$($_.pspath)\Volatile Environment") }).PSParentPath.split("\")[-1]

If ($NoSilent) {
	Set-Location "$Current_Folder\Sources\Run_in_Sandbox"
	powershell .\RunInSandbox_Config.ps1
}

$Sandbox_Icon = "$ProgData\Run_in_Sandbox\sandbox.ico"
$Run_in_Sandbox_Folder = "$ProgData\Run_in_Sandbox"
$XML_Config = "$Run_in_Sandbox_Folder\Sandbox_Config.xml"
$Get_XML_Content = [xml] (Get-Content $XML_Config)

# Check which context menu should be enabled
$Add_EXE = $Get_XML_Content.Configuration.ContextMenu_EXE
$Add_MSI = $Get_XML_Content.Configuration.ContextMenu_MSI
$Add_PS1 = $Get_XML_Content.Configuration.ContextMenu_PS1
$Add_VBS = $Get_XML_Content.Configuration.ContextMenu_VBS
$Add_ZIP = $Get_XML_Content.Configuration.ContextMenu_ZIP
$Add_Folder = $Get_XML_Content.Configuration.ContextMenu_Folder
$Add_Intunewin = $Get_XML_Content.Configuration.ContextMenu_Intunewin
$Add_MultipleApp = $Get_XML_Content.Configuration.ContextMenu_MultipleApp
$Add_Reg = $Get_XML_Content.Configuration.ContextMenu_Reg
$Add_ISO = $Get_XML_Content.Configuration.ContextMenu_ISO
$Add_PPKG = $Get_XML_Content.Configuration.ContextMenu_PPKG
$Add_HTML = $Get_XML_Content.Configuration.ContextMenu_HTML
$Add_MSIX = $Get_XML_Content.Configuration.ContextMenu_MSIX

If (-not (Test-Path "$ProgData\Run_in_Sandbox\RunInSandbox.ps1") ) {
	Write_Log -Message_Type "ERROR" -Message "File RunInSandbox.ps1 is missing"
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
	[System.Windows.Forms.MessageBox]::Show("File RunInSandbox.ps1 is missing !!!")
	break
}

$Backup_Folder = "$Destination_folder\Registry_Backup"
New-Item $Backup_Folder -Type Directory -Force | Out-Null
Write-Progress -Activity $Progress_Activity -PercentComplete 5

$List_Drive = Get-PSDrive | Where-Object { $_.Name -eq "HKCR_SD" }
If ($null -ne $List_Drive) { Remove-PSDrive $List_Drive }

Write_Log -Message_Type "INFO" -Message "Mapping registry HKCR"

$HKCR_Mapped = $False
Try {
	New-PSDrive -PSProvider registry -Root HKEY_CLASSES_ROOT -Name HKCR_SD | Out-Null
	Write_Log -Message_Type "SUCCESS" -Message "Mapping registry HKCR"
	$HKCR_Mapped = $True
}
Catch {
	Write_Log -Message_Type "ERROR" -Message "Mapping registry HKCR"
	Break
}

If ($HKCR_Mapped -ne $True) {
	Write_Log -Message_Type "ERROR" -Message "Could not map HKCR"
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
	[System.Windows.Forms.MessageBox]::Show("Could not map HKCR")
	break
}

Write-Progress -Activity $Progress_Activity -PercentComplete 10


Export_Reg_Config -Reg_Path "exefile" -Backup_Path "$Backup_Folder\Backup_HKRoot_EXEFile.reg"
Export_Reg_Config -Reg_Path "Microsoft.PowerShellScript.1" -Backup_Path "$Backup_Folder\Backup_HKRoot_PowerShellScript.reg"
Export_Reg_Config -Reg_Path "VBSFile" -Backup_Path "$Backup_Folder\Backup_HKRoot_VBSFile.reg"
Export_Reg_Config -Reg_Path "Msi.Package" -Backup_Path "$Backup_Folder\Backup_HKRoot_Msi.reg"
Export_Reg_Config -Reg_Path "CompressedFolder" -Backup_Path "$Backup_Folder\Backup_HKRoot_CompressedFolder.reg"
Export_Reg_Config -Reg_Path "WinRAR.ZIP" -Backup_Path "$Backup_Folder\Backup_HKRoot_WinRAR.reg"
Export_Reg_Config -Reg_Path "Directory" -Backup_Path "$Backup_Folder\Backup_HKRoot_Directory.reg"

Write-Progress -Activity $Progress_Activity -PercentComplete 15


Add-Content $log_file ""

Write_Log -Message_Type "INFO" -Message "Creating a restore point"
Checkpoint-Computer -Description "Add Windows Sandbox Context menus" -RestorePointType "MODIFY_SETTINGS" -ea silentlycontinue -ev ErrorRestore
If ($null -ne $ErrorRestore) {
	Write_Log -Message_Type "SUCCESS" -Message "Creation of restore point 'Add Windows Sandbox Context menus'"
}
Else {
	Write_Log -Message_Type "ERROR" -Message "Creation of restore point 'Add Windows Sandbox Context menus'"
	Write_Log -Message_Type "ERROR" -Message "$ErrorRestore"
}

Add-Content $log_file ""

Write-Progress -Activity $Progress_Activity -PercentComplete 20


If ($Add_PS1 -eq $True) {
	# ADD CONTEXT MENUS FOR PS1
	$PS1_Main_Menu = "Run PS1 in Sandbox"
	$PS1_SubMenu_RunAsUser = "Run PS1 as user"
	$PS1_SubMenu_RunAsSystem = "Run PS1 as system"
	$PS1_SubMenu_RunwithParams = "Run PS1 with parameters"

	$Command_For_Basic_PS1 = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -NoProfile -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type PS1Basic -LiteralPath "%V" -ScriptPath "%V"'
	$Command_For_Params_PS1 = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -NoProfile -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type PS1Params -LiteralPath "%V" -ScriptPath "%V"'
	$Command_For_System_PS1 = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -NoProfile -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type PS1System -LiteralPath "%V" -ScriptPath "%V"'

	$Windows_Version = (Get-CimInstance -class Win32_OperatingSystem).Caption
	If ($Windows_Version -like "*Windows 10*") {
		Write_Log -Message_Type "INFO" -Message "Running on Windows 10"

		$PS1_Shell_Registry_Key = "HKCR_SD:\SystemFileAssociations\.ps1\Shell"
		$Main_Menu_Path = "$PS1_Shell_Registry_Key\$PS1_Main_Menu"
		New-Item -Path $PS1_Shell_Registry_Key -Name $PS1_Main_Menu -Force | Out-Null
		New-ItemProperty -Path $Main_Menu_Path -Name "subcommands" -PropertyType String | Out-Null

		New-Item -Path $Main_Menu_Path -Name "Shell" -Force | Out-Null
		$Main_Menu_Shell_Path = "$Main_Menu_Path\Shell"

		New-Item -Path $Main_Menu_Shell_Path -Name $PS1_SubMenu_RunAsUser -Force | Out-Null
		New-Item -Path $Main_Menu_Shell_Path -Name $PS1_SubMenu_RunAsSystem -Force | Out-Null
		New-Item -Path $Main_Menu_Shell_Path -Name $PS1_SubMenu_RunwithParams -Force | Out-Null

		New-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsUser" -Name "Command" -Force | Out-Null
		New-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsSystem" -Name "Command" -Force | Out-Null
		New-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunwithParams" -Name "Command" -Force | Out-Null

		Set-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsUser\command" -Value $Command_For_Basic_PS1 -Force | Out-Null
		Set-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunwithParams\command" -Value $Command_For_Params_PS1 -Force | Out-Null
		Set-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsSystem\command" -Value $Command_For_System_PS1 -Force | Out-Null

		New-ItemProperty -Path "$Main_Menu_Path" -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
		New-ItemProperty -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunwithParams" -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
		New-ItemProperty -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsUser" -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
		New-ItemProperty -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsSystem" -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
	}
	If ($Windows_Version -like "*Windows 11*") {
		Write_Log -Message_Type "INFO" -Message "Running on Windows 11"
		$HKCU_Classes = "Registry::HKEY_USERS\$Current_User_SID" + "_Classes"
		If (Test-Path $HKCU_Classes) {
			$Default_PS1_HKCU = "$HKCU_Classes\.ps1"
			$rOpenWithProgids_Key = "$Default_PS1_HKCU\rOpenWithProgids"
			Write_Log -Message_Type "INFO" -Message "Checking programs from $rOpenWithProgids_Key"
			If (Test-Path $rOpenWithProgids_Key) {
				$Get_OpenWithProgids_Default_Value = (Get-Item $rOpenWithProgids_Key).Property
				ForEach ($Prop in $Get_OpenWithProgids_Default_Value) {
					$Default_HKCU_PS1_Shell_Registry_Key = "$HKCU_Classes\$Prop\Shell"
					If (Test-Path $Default_HKCU_PS1_Shell_Registry_Key) {
						$Main_Menu_Path = "$Default_HKCU_PS1_Shell_Registry_Key\$PS1_Main_Menu"

						Write_Log -Message_Type "INFO" -Message "Adding context menu for: $Main_Menu_Path"

						New-Item -Path $Default_HKCU_PS1_Shell_Registry_Key -Name $PS1_Main_Menu -Force | Out-Null
						New-ItemProperty -Path $Main_Menu_Path -Name "subcommands" -PropertyType String | Out-Null

						New-Item -Path $Main_Menu_Path -Name "Shell" -Force | Out-Null
						$Main_Menu_Shell_Path = "$Main_Menu_Path\Shell"

						New-Item -Path $Main_Menu_Shell_Path -Name $PS1_SubMenu_RunAsUser -Force | Out-Null
						New-Item -Path $Main_Menu_Shell_Path -Name $PS1_SubMenu_RunAsSystem -Force | Out-Null
						New-Item -Path $Main_Menu_Shell_Path -Name $PS1_SubMenu_RunwithParams -Force | Out-Null

						New-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsUser" -Name "Command" -Force | Out-Null
						New-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsSystem" -Name "Command" -Force | Out-Null
						New-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunwithParams" -Name "Command" -Force | Out-Null

						Set-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsUser\command" -Value $Command_For_Basic_PS1 -Force | Out-Null
						Set-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunwithParams\command" -Value $Command_For_Params_PS1 -Force | Out-Null
						Set-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsSystem\command" -Value $Command_For_System_PS1 -Force | Out-Null

						New-ItemProperty -Path "$Main_Menu_Path" -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
						New-ItemProperty -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunwithParams" -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
						New-ItemProperty -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsUser" -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
						New-ItemProperty -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsSystem" -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
					}
				}
			}

			$OpenWithProgids_Key = "$Default_PS1_HKCU\OpenWithProgids"
			If (Test-Path $OpenWithProgids_Key) {
				Write_Log -Message_Type "INFO" -Message "Checking programs from: $OpenWithProgids_Key"
				$Get_OpenWithProgids_Default_Value = (Get-Item $OpenWithProgids_Key).Property
				ForEach ($Prop in $Get_OpenWithProgids_Default_Value) {
					$Default_HKCU_PS1_Shell_Registry_Key = "$HKCU_Classes\$Prop\Shell"
					If (Test-Path $Default_HKCU_PS1_Shell_Registry_Key) {
						$Main_Menu_Path = "$Default_HKCU_PS1_Shell_Registry_Key\$PS1_Main_Menu"
						Write_Log -Message_Type "INFO" -Message "Adding context menu for: $Main_Menu_Path"
						New-Item -Path $Default_HKCU_PS1_Shell_Registry_Key -Name $PS1_Main_Menu -Force | Out-Null
						New-ItemProperty -Path $Main_Menu_Path -Name "subcommands" -PropertyType String | Out-Null

						New-Item -Path $Main_Menu_Path -Name "Shell" -Force | Out-Null
						$Main_Menu_Shell_Path = "$Main_Menu_Path\Shell"

						New-Item -Path $Main_Menu_Shell_Path -Name $PS1_SubMenu_RunAsUser -Force | Out-Null
						New-Item -Path $Main_Menu_Shell_Path -Name $PS1_SubMenu_RunAsSystem -Force | Out-Null
						New-Item -Path $Main_Menu_Shell_Path -Name $PS1_SubMenu_RunwithParams -Force | Out-Null

						New-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsUser" -Name "Command" -Force | Out-Null
						New-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsSystem" -Name "Command" -Force | Out-Null
						New-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunwithParams" -Name "Command" -Force | Out-Null

						Set-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsUser\command" -Value $Command_For_Basic_PS1 -Force | Out-Null
						Set-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunwithParams\command" -Value $Command_For_Params_PS1 -Force | Out-Null
						Set-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsSystem\command" -Value $Command_For_System_PS1 -Force | Out-Null

						New-ItemProperty -Path "$Main_Menu_Path" -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
						New-ItemProperty -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunwithParams" -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
						New-ItemProperty -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsUser" -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
						New-ItemProperty -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsSystem" -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
					}
				}
			}
		}


		# ADDING CONTEXT MENU DEPENDING OF THE USERCHOICE
		# The userchoice for PS1 is located in: HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.ps1\UserChoice
		# $Current_User_SID = (Get-ChildItem Registry::\HKEY_USERS | Where-Object { Test-Path "$($_.pspath)\Volatile Environment" } | ForEach-Object { (Get-ItemProperty "$($_.pspath)\Volatile Environment")}).PSParentPath.split("\")[-1]																			# RUN ON ISO
		$HKCU = "Registry::HKEY_USERS\$Current_User_SID"
		$HKCU_Classes = "Registry::HKEY_USERS\$Current_User_SID" + "_Classes"
		If (Test-Path $HKCU) {
			$PS1_UserChoice = "$HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.ps1\UserChoice"
			$Get_UserChoice = (Get-ItemProperty $PS1_UserChoice).ProgID

			Write_Log -Message_Type "INFO" -Message "Checking programs from: $PS1_UserChoice"

			$HKCR_UserChoice_Key = "HKCR_SD:\$Get_UserChoice"
			$HKCR_UserChoice_Shell = "$HKCR_UserChoice_Key\Shell"
			If (Test-Path $HKCR_UserChoice_Shell) {
				$HKCR_UserChoice_Label = "$HKCR_UserChoice_Shell\$PS1_Main_Menu"

				If (-not (Test-Path $HKCR_UserChoice_Label) ) {
					Write_Log -Message_Type "INFO" -Message "Adding context menu for: $HKCR_UserChoice_Label"
					New-Item -Path $HKCR_UserChoice_Label -Force | Out-Null
					New-ItemProperty -Path $HKCR_UserChoice_Label -Name "subcommands" -PropertyType String | Out-Null

					New-Item -Path $HKCR_UserChoice_Label -Name "Shell" -Force | Out-Null
					$Main_Menu_Shell_Path = "$HKCR_UserChoice_Label\Shell"

					New-Item -Path $Main_Menu_Shell_Path -Name $PS1_SubMenu_RunAsUser -Force | Out-Null
					New-Item -Path $Main_Menu_Shell_Path -Name $PS1_SubMenu_RunAsSystem -Force | Out-Null
					New-Item -Path $Main_Menu_Shell_Path -Name $PS1_SubMenu_RunwithParams -Force | Out-Null

					New-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsUser" -Name "Command" -Force | Out-Null
					New-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsSystem" -Name "Command" -Force | Out-Null
					New-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunwithParams" -Name "Command" -Force | Out-Null

					Set-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsUser\command" -Value $Command_For_Basic_PS1 -Force | Out-Null
					Set-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunwithParams\command" -Value $Command_For_Params_PS1 -Force | Out-Null
					Set-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsSystem\command" -Value $Command_For_System_PS1 -Force | Out-Null

					New-ItemProperty -Path "$HKCR_UserChoice_Label" -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
					New-ItemProperty -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunwithParams" -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
					New-ItemProperty -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsUser" -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
					New-ItemProperty -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsSystem" -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
				}
			}
		}
	}
	Write_Log -Message_Type "INFO" -Message "Context menus for PS1 have been added"
}

Write-Progress -Activity $Progress_Activity -PercentComplete 25


If ($Add_Intunewin -eq $True) {
	# RUN ON INTUNEWIN
	$Intunewin_Shell_Registry_Key = "HKCR_SD:\.intunewin"
	$Intunewin_Key_Label = "Test intunewin in Sandbox"
	$Intunewin_Key_Label_Path = "$Intunewin_Shell_Registry_Key\Shell\$Intunewin_Key_Label"
	$Intunewin_Command_Path = "$Intunewin_Key_Label_Path\Command"
	$Command_for_Intunewin = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -NoProfile -executionpolicy bypass -sta -windowstyle hidden -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type Intunewin -LiteralPath "%V" -ScriptPath "%V"'
	If (-not (Test-Path $Intunewin_Shell_Registry_Key) ) {
		New-Item $Intunewin_Shell_Registry_Key -Force | Out-Null
		New-Item "$Intunewin_Shell_Registry_Key\Shell" -Force | Out-Null
	}
	New-Item $Intunewin_Key_Label_Path -Force | Out-Null
	New-Item $Intunewin_Command_Path -Force | Out-Null
	# Set the command path
	Set-Item -Path $Intunewin_Command_Path -Value $Command_for_Intunewin -Force | Out-Null
	# Add Sandbox Icons
	New-ItemProperty -Path $Intunewin_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon -Force | Out-Null
	Write_Log -Message_Type "INFO" -Message "Context menus for IntuneWin have been added"

}

Write-Progress -Activity $Progress_Activity -PercentComplete 30


If ($Add_Reg -eq $True) {
	# RUN ON REG
	$Reg_Shell_Registry_Key = "HKCR_SD:\regfile\Shell"
	$Reg_Key_Label = "Test reg file in Sandbox"
	$Reg_Key_Label_Path = "$Reg_Shell_Registry_Key\$Reg_Key_Label"
	$Reg_Command_Path = "$Reg_Key_Label_Path\Command"
	$Command_for_Reg = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -NoProfile -executionpolicy bypass -sta -windowstyle hidden -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type REG -LiteralPath "%V" -ScriptPath "%V"'
	If (Test-Path $Reg_Shell_Registry_Key) {
		New-Item $Reg_Key_Label_Path | Out-Null
		New-Item $Reg_Command_Path | Out-Null
		# Set the command path
		Set-Item -Path $Reg_Command_Path -Value $Command_for_Reg -Force | Out-Null
		# Add Sandbox Icons
		New-ItemProperty -Path $Reg_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
		Write_Log -Message_Type "INFO" -Message "Context menus for REG have been added"
	}
}

Write-Progress -Activity $Progress_Activity -PercentComplete 35


If ($Add_ISO -eq $True) {
	# ADDING CONTEXT MENU FOR ISO
	$ISO_Key_Label = "Extract ISO file in Sandbox"
	$Command_for_ISO = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -NoProfile -executionpolicy bypass -sta -windowstyle hidden -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type ISO -LiteralPath "%V" -ScriptPath "%V"'

	Write_Log -Message_Type "INFO" -Message "Checking content of HKCR\.ISO"
	$ISO_Key = "HKCR_SD:\.ISO"
	If (Test-Path $ISO_Key) {
		Write_Log -Message_Type "INFO" -Message "The key HKCR\.ISO exists"
		$Get_ISO_Keys = Get-Item $ISO_Key
		ForEach ($Key in $Get_ISO_Keys) {
			$Get_Properties = $Key.Property
			Write_Log -Message_Type "INFO" -Message "Following subkeys found: $Get_Properties"
			foreach ($Property in $Get_Properties) {
				$Prop = (Get-ItemProperty $ISO_Key)."$Property"
				Write_Log -Message_Type "INFO" -Message "Following property found: $Prop"
				$ISO_Property_Key = "HKCR_SD:\$Prop"
				Write_Log -Message_Type "INFO" -Message "Reg path to test: $ISO_Property_Key"
				If (Test-Path $ISO_Property_Key) {
					Write_Log -Message_Type "INFO" -Message "The following reg path exists: $ISO_Property_Key"
					$ISO_Property_Shell = "$ISO_Property_Key\Shell"
					If (-not (Test-Path $ISO_Property_Shell) ) {
						New-Item $ISO_Property_Shell | Out-Null
					}
					$ISO_Key_Label_Path = "$ISO_Property_Shell\$ISO_Key_Label"
					If (-not (Test-Path $ISO_Key_Label_Path) ) {
						$ISO_Command_Path = "$ISO_Key_Label_Path\Command"
						New-Item $ISO_Key_Label_Path | Out-Null
						New-Item $ISO_Command_Path | Out-Null
						# Set the command path
						Set-Item -Path $ISO_Command_Path -Value $Command_for_ISO -Force | Out-Null
						# Add Sandbox Icons
						New-ItemProperty -Path $ISO_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
						Write_Log -Message_Type "INFO" -Message "Creating following context menu for ISO under: $ISO_Key_Label_Path"
					}
				}
				Else {
					Write_Log -Message_Type "INFO" -Message "The following reg path does not exist: $ISO_Property_Shell"
				}
			}
		}
	}

	# Modify value from HKCU
	$HKCU_Classes = "Registry::HKEY_USERS\$Current_User_SID" + "_Classes"
	If (Test-Path $HKCU_Classes) {
		$Default_ISO_HKCU = "$HKCU_Classes\.iso"
		If (Test-Path $Default_ISO_HKCU) {
			$Get_Default_Value = (Get-ItemProperty $Default_ISO_HKCU)."(default)"
			$Default_HKCU_ISO_Shell_Registry_Key = "$HKCU_Classes\$Get_Default_Value\Shell"
			If (Test-Path $Default_HKCU_ISO_Shell_Registry_Key) {
				$ISO_HKCU_Key_Label_Path = "$Default_HKCU_ISO_Shell_Registry_Key\$ISO_Key_Label"
				If (-not (Test-Path $ISO_HKCU_Key_Label_Path) ) {
					$HKCU_ISO_Command_Path = "$ISO_HKCU_Key_Label_Path\Command"
					New-Item $ISO_HKCU_Key_Label_Path | Out-Null
					New-Item $HKCU_ISO_Command_Path | Out-Null
					# Set the command path
					Set-Item -Path $HKCU_ISO_Command_Path -Value $Command_for_ISO -Force | Out-Null
					# Add Sandbox Icons
					New-ItemProperty -Path $ISO_HKCU_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
					Write_Log -Message_Type "INFO" -Message "Context menu for ISO has been added"
				}
			}
		}
	}
}

Write-Progress -Activity $Progress_Activity -PercentComplete 40


If ($Add_PPKG -eq $True) {
	# RUN ON PPKG
	$PPKG_Shell_Registry_Key = "HKCR_SD:\Microsoft.ProvTool.Provisioning.1\Shell"
	$PPKG_Key_Label = "Run PPKG file in Sandbox"
	$PPKG_Key_Label_Path = "$PPKG_Shell_Registry_Key\$PPKG_Key_Label"
	$PPKG_Command_Path = "$PPKG_Key_Label_Path\Command"
	$Command_for_PPKG = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -NoProfile -executionpolicy bypass -sta -windowstyle hidden -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type PPKG -LiteralPath "%V" -ScriptPath "%V"'
	If (Test-Path $PPKG_Shell_Registry_Key) {
		New-Item $PPKG_Key_Label_Path | Out-Null
		New-Item $PPKG_Command_Path | Out-Null
		# Set the command path
		Set-Item -Path $PPKG_Command_Path -Value $Command_for_PPKG -Force | Out-Null
		# Add Sandbox Icons
		New-ItemProperty -Path $PPKG_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
		Write_Log -Message_Type "INFO" -Message "Context menu for PPKG has been added"
	}
}

Write-Progress -Activity $Progress_Activity -PercentComplete 40


If ($Add_HTML -eq $True) {
	$HTML_Key_Label = "Run this web link in Sandbox"

	# RUN ON HTML for Edge
	$HTML_Edge_Shell_Registry_Key = "HKCR_SD:\MSEdgeHTM\Shell"
	If (Test-Path $HTML_Edge_Shell_Registry_Key) {
		$HTML_Edge_Key_Label_Path = "$HTML_Edge_Shell_Registry_Key\$HTML_Key_Label"
		If (-not (Test-Path $HTML_Edge_Key_Label_Path) ) {
			$HTML_Edge_Command_Path = "$HTML_Edge_Key_Label_Path\Command"
			$Command_for_HTML = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -NoProfile -executionpolicy bypass -sta -windowstyle hidden -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type HTML -LiteralPath "%V" -ScriptPath "%V"'
			New-Item $HTML_Edge_Key_Label_Path | Out-Null
			New-Item $HTML_Edge_Command_Path | Out-Null
			# Set the command path
			Set-Item -Path $HTML_Edge_Command_Path -Value $Command_for_HTML -Force | Out-Null
			# Add Sandbox Icons
			New-ItemProperty -Path $HTML_Edge_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
			Write_Log -Message_Type "INFO" -Message "Context menu for HTML has been added"
		}
	}

	# RUN ON HTML for Chrome
	$HTML_Chrome_Shell_Registry_Key = "HKCR_SD:\ChromeHTML\Shell"
	If (Test-Path $HTML_Chrome_Shell_Registry_Key) {
		$HTML_Chrome_Key_Label_Path = "$HTML_Chrome_Shell_Registry_Key\$HTML_Key_Label"
		If (-not (Test-Path $HTML_Chrome_Key_Label_Path) ) {
			$HTML_Chrome_Command_Path = "$HTML_Chrome_Key_Label_Path\Command"
			$Command_for_HTML = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -NoProfile -executionpolicy bypass -sta -windowstyle hidden -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type HTML -LiteralPath "%V" -ScriptPath "%V"'

			New-Item $HTML_Chrome_Key_Label_Path | Out-Null
			New-Item $HTML_Chrome_Command_Path | Out-Null
			# Set the command path
			Set-Item -Path $HTML_Chrome_Command_Path -Value $Command_for_HTML -Force | Out-Null
			# Add Sandbox Icons
			New-ItemProperty -Path $HTML_Chrome_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
			Write_Log -Message_Type "INFO" -Message "Context menu for HTML has been added"
		}
	}

	# RUN ON HTML for IE
	$HTML_IE_Shell_Registry_Key = "HKCR_SD:\IE.AssocFile.HTM\Shell"
	If (Test-Path $HTML_IE_Shell_Registry_Key) {
		$HTML_IE_Key_Label_Path = "$HTML_IE_Shell_Registry_Key\$HTML_Key_Label"
		If (-not (Test-Path $HTML_IE_Key_Label_Path) ) {
			$HTML_IE_Command_Path = "$HTML_IE_Key_Label_Path\Command"
			$Command_for_HTML = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -NoProfile -executionpolicy bypass -sta -windowstyle hidden -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type HTML -LiteralPath "%V" -ScriptPath "%V"'
			New-Item $HTML_IE_Key_Label_Path | Out-Null
			New-Item $HTML_IE_Command_Path | Out-Null
			# Set the command path
			Set-Item -Path $HTML_IE_Command_Path -Value $Command_for_HTML -Force | Out-Null
			# Add Sandbox Icons
			New-ItemProperty -Path $HTML_IE_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
			Write_Log -Message_Type "INFO" -Message "Context menu for HTML has been added"
		}
	}

	$URL_Shell_Registry_Key = "HKCR_SD:\IE.AssocFile.URL\Shell"
	If (Test-Path $URL_Shell_Registry_Key) {
		$URL_Key_Label_Path = "$URL_Shell_Registry_Key\Run this URL in Sandbox"
		If (-not (Test-Path $URL_Key_Label_Path) ) {
			$URL_Command_Path = "$URL_Key_Label_Path\Command"
			$Command_for_URL = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -NoProfile -executionpolicy bypass -sta -windowstyle hidden -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type URL -LiteralPath "%V" -ScriptPath "%V"'
			New-Item $URL_Key_Label_Path | Out-Null
			New-Item $URL_Command_Path | Out-Null
			# Set the command path
			Set-Item -Path $URL_Command_Path -Value $Command_for_URL -Force | Out-Null
			# Add Sandbox Icons
			New-ItemProperty -Path $URL_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
			Write_Log -Message_Type "INFO" -Message "Context menu for URL has been added"
		}
	}


	# ADD CONTEXT MENU FOR HTML ISING USERCHOICE
	# $Current_User_SID = (Get-ChildItem Registry::\HKEY_USERS | Where-Object { Test-Path "$($_.pspath)\Volatile Environment" } | ForEach-Object { (Get-ItemProperty "$($_.pspath)\Volatile Environment")}).PSParentPath.split("\")[-1]																			# RUN ON ISO
	$HKCU = "Registry::HKEY_USERS\$Current_User_SID"
	$HKCU_Classes = "Registry::HKEY_USERS\$Current_User_SID" + "_Classes"
	If (Test-Path $HKCU) {
		$HTML_UserChoice = "$HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.html\UserChoice"
		$Get_UserChoice = (Get-ItemProperty $HTML_UserChoice).ProgID
		$HKCR_UserChoice_Key = "HKCR_SD:\$Get_UserChoice"
		$HKCR_UserChoice_Shell = "$HKCR_UserChoice_Key\Shell"
		If (Test-Path $HKCR_UserChoice_Shell) {
			$HKCR_UserChoice_Label = "$HKCR_UserChoice_Shell\$HTML_Key_Label"
			If (-not (Test-Path $HKCR_UserChoice_Label) ) {
				$HTML_UserChoice_Command_Path = "$HKCR_UserChoice_Label\Command"
				$Command_for_HTML = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -NoProfile -executionpolicy bypass -sta -windowstyle hidden -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type HTML -LiteralPath "%V" -ScriptPath "%V"'
				New-Item $HKCR_UserChoice_Label | Out-Null
				New-Item $HTML_UserChoice_Command_Path | Out-Null
				# Set the command path
				Set-Item -Path $HTML_UserChoice_Command_Path -Value $Command_for_HTML -Force | Out-Null
				# Add Sandbox Icons
				New-ItemProperty -Path $HKCR_UserChoice_Label -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
				Write_Log -Message_Type "INFO" -Message "Context menu for HTML has been added"
			}
		}
	}
}

Write-Progress -Activity $Progress_Activity -PercentComplete 45

If ($Add_MultipleApp -eq $True) {
	# RUN ON bundle app
	$MultipleApps_Shell_Registry_Key = "HKCR_SD:\.sdbapp"
	$MultipleApps_Key_Label = "Test application bundle in Sandbox"
	$MultipleApps_Key_Label_Path = "$MultipleApps_Shell_Registry_Key\Shell\$MultipleApps_Key_Label"
	$MultipleApps_Command_Path = "$MultipleApps_Key_Label_Path\Command"
	$Command_for_MultipleApps = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -NoProfile -executionpolicy bypass -sta -windowstyle hidden -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type SDBApp -LiteralPath "%V" -ScriptPath "%V"'
	If (-not (Test-Path $MultipleApps_Shell_Registry_Key) ) {
		New-Item $MultipleApps_Shell_Registry_Key | Out-Null
		New-Item "$MultipleApps_Shell_Registry_Key\Shell" | Out-Null
		New-Item $MultipleApps_Key_Label_Path | Out-Null
		New-Item $MultipleApps_Command_Path | Out-Null
		# Set the command path
		Set-Item -Path $MultipleApps_Command_Path -Value $Command_for_MultipleApps -Force | Out-Null
		# Add Sandbox Icons
		New-ItemProperty -Path $MultipleApps_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
		Write_Log -Message_Type "INFO" -Message "Context menu for PS1 has been added"
	}
}

Write-Progress -Activity $Progress_Activity -PercentComplete 45

If ($Add_VBS -eq $True) {
	# RUN ON VBS
	$VBS_Shell_Registry_Key = "HKCR_SD:\VBSFile\Shell"
	$VBS_Basic_Run = "Run VBS in Sandbox"
	$VBS_Parameter_Run = "Run VBS in Sandbox with parameters"

	$ContextMenu_Basic_VBS = "$VBS_Shell_Registry_Key\$VBS_Basic_Run"
	$ContextMenu_Parameters_VBS = "$VBS_Shell_Registry_Key\$VBS_Parameter_Run"

	New-Item -Path $VBS_Shell_Registry_Key -Name $VBS_Basic_Run -Force | Out-Null
	New-Item -Path $VBS_Shell_Registry_Key -Name $VBS_Parameter_Run -Force | Out-Null
	New-Item -Path $ContextMenu_Basic_VBS -Name "Command" -Force | Out-Null
	New-Item -Path $ContextMenu_Parameters_VBS -Name "Command" -Force | Out-Null
	$Command_For_Basic_VBS = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -NoProfile -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type VBSBasic -LiteralPath "%V" -ScriptPath "%V"'
	$Command_For_Params_VBS = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -NoProfile -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type VBSParams -LiteralPath "%V" -ScriptPath "%V"'
	# Set the command path
	Set-Item -Path "$ContextMenu_Basic_VBS\command" -Value $Command_For_Basic_VBS -Force | Out-Null
	Set-Item -Path "$ContextMenu_Parameters_VBS\command" -Value $Command_For_Params_VBS -Force | Out-Null
	# Add Sandbox Icons
	New-ItemProperty -Path $ContextMenu_Basic_VBS -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
	New-ItemProperty -Path $ContextMenu_Parameters_VBS -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
	Write_Log -Message_Type "INFO" -Message "Context menus for VBS have been added"
}

Write-Progress -Activity $Progress_Activity -PercentComplete 50

If ($Add_EXE -eq $True) {
	# RUN ON EXE
	$EXE_Shell_Registry_Key = "HKCR_SD:\exefile\Shell"
	$EXE_Basic_Run = "Run EXE in Sandbox"
	$ContextMenu_Basic_EXE = "$EXE_Shell_Registry_Key\$EXE_Basic_Run"

	New-Item -Path $EXE_Shell_Registry_Key -Name $EXE_Basic_Run -Force | Out-Null
	New-Item -Path $ContextMenu_Basic_EXE -Name "Command" -Force | Out-Null
	# Add Sandbox Icons
	New-ItemProperty -Path $ContextMenu_Basic_EXE -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
	$Command_For_EXE = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -NoProfile -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type EXE -LiteralPath "%V" -ScriptPath "%V"'
	# Set the command path
	Set-Item -Path "$ContextMenu_Basic_EXE\command" -Value $Command_For_EXE -Force | Out-Null
	Write_Log -Message_Type "INFO" -Message "Context menus for EXE have been added"
}

Write-Progress -Activity $Progress_Activity -PercentComplete 50

If ($Add_MSI -eq $True) {
	# RUN ON MSI
	$MSI_Shell_Registry_Key = "HKCR_SD:\Msi.Package\Shell"
	$MSI_Basic_Run = "Run MSI in Sandbox"
	$ContextMenu_Basic_MSI = "$MSI_Shell_Registry_Key\$MSI_Basic_Run"

	New-Item -Path $MSI_Shell_Registry_Key -Name $MSI_Basic_Run -Force | Out-Null
	New-Item -Path $ContextMenu_Basic_MSI -Name "Command" -Force | Out-Null
	# Add Sandbox Icons
	New-ItemProperty -Path $ContextMenu_Basic_MSI -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
	$Command_For_MSI = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -NoProfile -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type MSI -LiteralPath "%V" -ScriptPath "%V"'
	# Set the command path
	Set-Item -Path "$ContextMenu_Basic_MSI\command" -Value $Command_For_MSI -Force | Out-Null
	Write_Log -Message_Type "INFO" -Message "Context menu for MSI has been added"
}

Write-Progress -Activity $Progress_Activity -PercentComplete 55


If ($Add_ZIP -eq $True) {
	# RUN ON ZIP
	$ZIP_Shell_Registry_Key = "HKCR_SD:\CompressedFolder\Shell"
	$ZIP_Basic_Run = "Extract ZIP in Sandbox"
	$ContextMenu_Basic_ZIP = "$ZIP_Shell_Registry_Key\$ZIP_Basic_Run"

	New-Item -Path $ZIP_Shell_Registry_Key -Name $ZiP_Basic_Run -Force | Out-Null
	New-Item -Path $ContextMenu_Basic_ZIP -Name "Command" -Force | Out-Null
	# Add Sandbox Icons
	New-ItemProperty -Path $ContextMenu_Basic_ZIP -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
	$Command_For_ZIP = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -NoProfile -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type ZIP -LiteralPath "%V" -ScriptPath "%V"'
	# Set the command path
	Set-Item -Path "$ContextMenu_Basic_ZIP\command" -Value $Command_For_ZIP -Force | Out-Null
	Write_Log -Message_Type "INFO" -Message "Context menu for ZIP has been added"

	Write-Progress -Activity $Progress_Activity -PercentComplete 65

	# RUN ON ZIP if WinRAR is installed
	If (Test-Path "HKCR_SD:\WinRAR.ZIP\Shell") {
		$ZIP_WinRAR_Shell_Registry_Key = "HKCR_SD:\WinRAR.ZIP\Shell"
		$ZIP_WinRAR_Basic_Run = "Extract ZIP in Sandbox"
		$ContextMenu_Basic_ZIP_RAR = "$ZIP_WinRAR_Shell_Registry_Key\$ZIP_WinRAR_Basic_Run"

		New-Item -Path $ZIP_WinRAR_Shell_Registry_Key -Name $ZIP_WinRAR_Basic_Run -Force | Out-Null
		New-Item -Path $ContextMenu_Basic_ZIP_RAR -Name "Command" -Force | Out-Null
		# Add Sandbox Icons
		New-ItemProperty -Path $ContextMenu_Basic_ZIP_RAR -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
		$Command_For_ZIP = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -NoProfile -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type ZIP -LiteralPath "%V" -ScriptPath "%V"'
		# Set the command path
		Set-Item -Path "$ContextMenu_Basic_ZIP_RAR\command" -Value $Command_For_ZIP -Force | Out-Null
	}

	$7z_Key_Label = "Extract 7z file in Sandbox"
	# RUN ON 7z
	$7z_Shell_Registry_Key = "HKCR_SD:\.7z"
	If (Test-Path $7z_Shell_Registry_Key) {
		$Get_Default_Value = (Get-ItemProperty "HKCR_SD:\.7z")."(default)"
		$Default_ZIP_Shell_Registry_Key = "HKCR_SD:\$Get_Default_Value\Shell"
		If (Test-Path $Default_ZIP_Shell_Registry_Key) {
			$Default_ZIP_Key_Label_Path = "$Default_ZIP_Shell_Registry_Key\$7z_Key_Label"
			$Default_ZIP_Command_Path = "$Default_ZIP_Key_Label_Path\Command"
			$Command_for_Default_ZIP = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -NoProfile -executionpolicy bypass -sta -windowstyle hidden -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type 7Z -LiteralPath "%V" -ScriptPath "%V"'
			If (Test-Path $Default_ZIP_Shell_Registry_Key) {
				New-Item $Default_ZIP_Key_Label_Path | Out-Null
				New-Item $Default_ZIP_Command_Path | Out-Null
				# Set the command path
				Set-Item -Path $Default_ZIP_Command_Path -Value $Command_for_Default_ZIP -Force | Out-Null
				# Add Sandbox Icons
				New-ItemProperty -Path $Default_ZIP_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
				Write_Log -Message_Type "INFO" -Message "Context menu for 7Z has been added"
			}
		}
	}

	# $Current_User_SID = (Get-ChildItem Registry::\HKEY_USERS | Where-Object { Test-Path "$($_.pspath)\Volatile Environment" } | ForEach-Object { (Get-ItemProperty "$($_.pspath)\Volatile Environment")}).PSParentPath.split("\")[-1]																			# RUN ON ISO
	$HKCU = "Registry::HKEY_USERS\$Current_User_SID"
	$HKCU_Classes = "Registry::HKEY_USERS\$Current_User_SID" + "_Classes"
	If (Test-Path $HKCU) {
		$ZIP_UserChoice = "$HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\OpenWithProgids"
		$Get_Properties = (Get-Item $ZIP_UserChoice).property
		ForEach ($Prop in $Get_Properties) {
			$HKCR_UserChoice_Key = "HKCR_SD:\$Prop"
			$HKCR_UserChoice_Key
			$HKCR_UserChoice_Shell = "$HKCR_UserChoice_Key\Shell"
			If (Test-Path $HKCR_UserChoice_Shell) {
				$HKCR_UserChoice_Label = "$HKCR_UserChoice_Shell\$ZIP_Basic_Run"
				If (-not (Test-Path $HKCR_UserChoice_Label) ) {
					$HTML_UserChoice_Command_Path = "$HKCR_UserChoice_Label\Command"
					$Command_for_HTML = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -NoProfile -executionpolicy bypass -sta -windowstyle hidden -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type HTML -LiteralPath "%V" -ScriptPath "%V"'
					New-Item $HKCR_UserChoice_Label | Out-Null
					New-Item $HTML_UserChoice_Command_Path | Out-Null
					# Set the command path
					Set-Item -Path $HTML_UserChoice_Command_Path -Value $Command_for_HTML -Force | Out-Null
					# Add Sandbox Icons
					New-ItemProperty -Path $HKCR_UserChoice_Label -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
					Write_Log -Message_Type "INFO" -Message "Context menu for HTML has been added"
				}
			}
		}
	}
}

Write-Progress -Activity $Progress_Activity -PercentComplete 75

# RUN ON MSIX
If ($Add_MSIX -eq $True) {
	$MSIX_Key_Label = "Run MSIX file in Sandbox"
	$Command_for_MSIX = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -NoProfile -executionpolicy bypass -sta -windowstyle hidden -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type MSIX -LiteralPath "%V" -ScriptPath "%V"'

	$MSIX_Shell_Registry_Key = "HKCR_SD:\.msix\OpenWithProgids"
	If (Test-Path $MSIX_Shell_Registry_Key) {
		$Get_Default_Value = (Get-Item $MSIX_Shell_Registry_Key).Property
		$MSIX_Shell_Registry = "HKCR_SD:\$Get_Default_Value\Shell"
		If (Test-Path $MSIX_Shell_Registry) {
			$MSIX_Key_Label_Path = "$MSIX_Shell_Registry\$MSIX_Key_Label"
			If (-not (Test-Path $MSIX_Key_Label_Path) ) {
				$MSIX_Command_Path = "$MSIX_Key_Label_Path\Command"
				New-Item $MSIX_Key_Label_Path | Out-Null
				New-Item $MSIX_Command_Path | Out-Null
				# Set the command path
				Set-Item -Path $MSIX_Command_Path -Value $Command_for_MSIX -Force | Out-Null
				# Add Sandbox Icons
				New-ItemProperty -Path $MSIX_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
				Write_Log -Message_Type "INFO" -Message "Context menu for MSIX has been added"
			}
		}
	}

	# Modify value from HKCU
	# $Current_User_SID = (Get-ChildItem Registry::\HKEY_USERS | Where-Object { Test-Path "$($_.pspath)\Volatile Environment" } | ForEach-Object { (Get-ItemProperty "$($_.pspath)\Volatile Environment")}).PSParentPath.split("\")[-1]																			# RUN ON ISO
	$HKCU_Classes = "Registry::HKEY_USERS\$Current_User_SID" + "_Classes"
	If (Test-Path $HKCU_Classes) {
		$Default_MSIX_HKCU = "$HKCU_Classes\.msix"
		$Get_Default_Value = (Get-Item "$Default_MSIX_HKCU\OpenWithProgids").Property
		$Default_HKCU_MSIX_Shell_Registry_Key = "$HKCU_Classes\$Get_Default_Value\Shell"
		If (Test-Path $Default_HKCU_MSIX_Shell_Registry_Key) {
			$MSIX_HKCU_Key_Label_Path = "$Default_HKCU_MSIX_Shell_Registry_Key\$MSIX_Key_Label"
			$HKCU_MSIX_Command_Path = "$MSIX_HKCU_Key_Label_Path\Command"
			New-Item $MSIX_HKCU_Key_Label_Path | Out-Null
			New-Item $HKCU_MSIX_Command_Path | Out-Null
			# Set the command path
			Set-Item -Path $HKCU_MSIX_Command_Path -Value $Command_for_MSIX -Force | Out-Null
			# Add Sandbox Icons
			New-ItemProperty -Path $MSIX_HKCU_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
			Write_Log -Message_Type "INFO" -Message "Context menu for ISO has been added"
		}
	}
}

Write-Progress -Activity $Progress_Activity -PercentComplete 75


If ($Add_Folder -eq $True) {
	# Share this folder - Inside the folder
	$Folder_Inside_Shell_Registry_Key = "HKCR_SD:\Directory\Background\shell"
	$Folder_Inside_Basic_Run = "Share this folder in a Sandbox"
	# $Folder_Inside_Basic_Run = $Get_Language_File_Content.Folder
	# $ContextMenu_Folder_Inside = "$Folder_Inside_Shell_Registry_Key\$Folder_Inside_Basic_Run"
	$ContextMenu_Folder_Inside = "$Folder_Inside_Shell_Registry_Key\Share this folder in a Sandbox"

	New-Item -Path $Folder_Inside_Shell_Registry_Key -Name $Folder_Inside_Basic_Run -Force | Out-Null
	New-Item -Path $ContextMenu_Folder_Inside -Name "Command" -Force | Out-Null
	# Add Sandbox Icons
	New-ItemProperty -Path $ContextMenu_Folder_Inside -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
	$Command_For_Folder_Inside = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -NoProfile -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type Folder_Inside -LiteralPath "%V" -ScriptPath "%V"'
	# Set the command path
	Set-Item -Path "$ContextMenu_Folder_Inside\command" -Value $Command_For_Folder_Inside -Force | Out-Null
	Write_Log -Message_Type "INFO" -Message "Context menus for folder have been added"


	Write-Progress -Activity $Progress_Activity -PercentComplete 85


	# Share this folder - Right-click on the folder
	$Folder_On_Shell_Registry_Key = "HKCR_SD:\Directory\shell"
	# $Folder_On_Run = $Get_Language_File_Content.Folder
	$Folder_On_Run = "Share this folder in a Sandbox"
	# $ContextMenu_Folder_On = "$Folder_On_Shell_Registry_Key\$Folder_On_Run"
	$ContextMenu_Folder_On = "$Folder_On_Shell_Registry_Key\Share this folder in a Sandbox"

	New-Item -Path $Folder_On_Shell_Registry_Key -Name $Folder_On_Run -Force | Out-Null
	New-Item -Path $ContextMenu_Folder_On -Name "Command" -Force | Out-Null
	# Add Sandbox Icons
	New-ItemProperty -Path $ContextMenu_Folder_On -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
	$Command_For_Folder_On = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -NoProfile -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type Folder_On -LiteralPath "%V" -ScriptPath "%V"'
	# Set the command path
	Set-Item -Path "$ContextMenu_Folder_On\command" -Value $Command_For_Folder_On -Force | Out-Null
}

If ($null -ne $List_Drive) { Remove-PSDrive $List_Drive }

Write-Progress -Activity $Progress_Activity -PercentComplete 100
Copy-Item $Log_File $Destination_folder -Force
