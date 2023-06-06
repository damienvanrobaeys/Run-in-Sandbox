<#
# Author & creator: Damien VAN ROBAEYS
# Website: http://www.systanddeploy.com
# Twitter: https://twitter.com/syst_and_deploy

Contributor: Joly0 with below GitHub PR
- Added option to run cmd/bat files in sandbox (solves Run CMD/BAT as user or system in Sandbox #21)
- Added option to run pdf-files in sandbox (these should be covered by run in html, but does not, if another program is default for pdf, other than edge/chrome/etc)
- Added option to cleanup wsb file after closing the sandbox (solves Trash wbs file after closing sandbox #4)
- Completly rewrote Add_Structure.ps1 for better readability and expansion in further releases
- Outsourced changelog to separate changelog.md
- Added ServiceUI in favor of psexec
- Fixed a lot of issues with various context menu´s not correctly working/being added

Contributor: ImportTaste with below GitHub PR
Add a switch to skip checkpoint creation
Add PSEdition Desktop requirement

Contributor: Harm Veenstra with below GitHub PR
Formatting and noprofile addition to all powershell commands being started
#>

#Requires -PSEdition Desktop

Param
(
	[Switch]$NoSilent,
	[Switch]$NoCheckpoint	
)

$TEMP_Folder = $env:temp
$Log_File = "$TEMP_Folder\RunInSandbox_Install.log"
$Current_Folder = Split-Path $MyInvocation.MyCommand.Path

$Script:Current_User_SID = (Get-ChildItem Registry::\HKEY_USERS | Where-Object { Test-Path "$($_.pspath)\Volatile Environment" } | ForEach-Object { (Get-ItemProperty "$($_.pspath)\Volatile Environment") }).PSParentPath.split("\")[-1]
$HKCU = "Registry::HKEY_USERS\$Current_User_SID"
$HKCU_Classes = "Registry::HKEY_USERS\$Current_User_SID" + "_Classes"
$Windows_Version = (Get-CimInstance -class Win32_OperatingSystem).Caption

If(test-path $Log_File){remove-item $Log_File}
new-item $Log_File -type file -force | out-null

Function Write_Log
	{
		param(
		$Message_Type,	
		$Message
		)
		
		$MyDate = "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)		
		Add-Content $Log_File  "$MyDate - $Message_Type : $Message"	
		write-host "$MyDate - $Message_Type : $Message"	
	}
	
Function Export_Reg_Config
	{
		param(
		$Reg_Path,
		$Backup_Path
		)
						
		If (Test-Path "Registry::HKEY_CLASSES_ROOT\$Reg_Path") 
			{
				Try
					{
						reg export "HKEY_CLASSES_ROOT\$Reg_Path" $Backup_Path /y | Out-Null
						Write_Log -Message_Type "SUCCESS" -Message "$Reg_Path has been exported"									
					}
				Catch
					{
						Write_Log -Message_Type "ERROR" -Message "$Reg_Path has not been exported"									
					}			
			}
		Else
			{
				Write_Log -Message_Type "INFO" -Message "Can not find registry path: Registry::HKEY_CLASSES_ROOT\$Reg_Path"
			}				
		Add-content $log_file ""		
	}	

	
Write_Log -Message_Type "INFO" -Message "Starting the configuration of RunInSandbox"												

[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | out-null
$Run_As_Admin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
If($Run_As_Admin -eq $False)
	{
		Write_Log -Message_Type "ERROR" -Message "The script has not been lauched with admin rights"													
		[System.Windows.Forms.MessageBox]::Show("Please run the tool with admin rights :-)")
		break		
	}

Write_Log -Message_Type "SUCCESS" -Message "The script has been launched with admin rights"

$Is_Sandbox_Installed = (Get-WindowsOptionalFeature -online | where {$_.featurename -eq "Containers-DisposableClientVM"}).state
If ($Is_Sandbox_Installed -eq "Disabled") 
	{
		Write-Log -Message_Type "ERROR" -Message "The feature `"Windows Sandbox`" is not installed !!!"
		[System.Windows.Forms.MessageBox]::Show("The feature `"Windows Sandbox`" is not installed !!!")
		break
	}		

$Current_Folder = split-path $MyInvocation.MyCommand.Path
$Sources = $Current_Folder + "\" + "Sources\*"			
If(-not (Test-Path $Sources)) 
	{
		Write_Log -Message_Type "ERROR" -Message "Sources folder is missing"
		[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
		[System.Windows.Forms.MessageBox]::Show("It seems you haven´t downloaded all the folder structure.`nThe folder `"Sources`" is missing !!!")
		break
	}				
				
Write_Log -Message_Type "SUCCESS" -Message "The sources folder exists"		

Add-content $log_file ""		

$Progress_Activity = "Enabling Run in Sandbox context menus"
write-progress -activity $Progress_Activity -percentcomplete 1;
		
$Check_Sources_Files_Count = (get-childitem "$Current_Folder\Sources\Run_in_Sandbox" -recurse).count

If($Check_Sources_Files_Count -ne 40) 
	{
		Write_Log -Message_Type "ERROR" -Message "Some contents are missing"
		[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
		[System.Windows.Forms.MessageBox]::Show("It seems you haven´t downloaded all the folder structure !!!")
		break
	}				

$Sources_Copied = $False
$Destination_folder = "$env:ProgramData\Run_in_Sandbox"
Try 
	{
		Copy-Item $Sources $env:ProgramData -Force -Recurse | Out-Null
		Write_Log -Message_Type "SUCCESS" -Message "Sources have been copied in $env:ProgramData\Run_in_Sandbox"
	}
Catch 
	{
		Write_Log -Message_Type "ERROR" -Message "Sources have not been copied in $env:ProgramData\Run_in_Sandbox"
		EXIT
	}				

$Sources_Unblocked = $False
Try 
	{
		Get-ChildItem -Recurse $Destination_folder | Unblock-File
		Write_Log -Message_Type "SUCCESS" -Message "Sources files have been unblocked"
		$Sources_Unblocked = $True
	}
Catch 
	{
		Write-Log -Message_Type "ERROR" -Message "Sources files have not been unblocked"
		break
	}

If($Sources_Unblocked -ne $True) 
	{
		Write_Log -Message_Type "ERROR" -Message "Source files could not be unblocked"
		[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
		[System.Windows.Forms.MessageBox]::Show("Source files could not be unblocked")
		break
	}					

If($NoSilent)
	{
		Set-Location "$Current_Folder\Sources\Run_in_Sandbox"
		powershell -NoProfile  .\RunInSandbox_Config.ps1
	}

$Sandbox_Icon = "$env:ProgramData\Run_in_Sandbox\sandbox.ico"
$Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"
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
$Add_CMD = $Get_XML_Content.Configuration.ContextMenu_CMD
$Add_PDF = $Get_XML_Content.Configuration.ContextMenu_PDF				

If(-not (Test-Path "$env:ProgramData\Run_in_Sandbox\RunInSandbox.ps1") ) {
	Write_Log -Message_Type "ERROR" -Message "File RunInSandbox.ps1 is missing"
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
	[System.Windows.Forms.MessageBox]::Show("File RunInSandbox.ps1 is missing !!!")
	break
}

$Backup_Folder = "$Destination_folder\Registry_Backup"
new-item $Backup_Folder -Type Directory -Force | out-null

write-progress -activity $Progress_Activity  -percentcomplete 5;

Write_Log -Message_Type "INFO" -Message "Exporting registry keys"		

Export_Reg_Config -Reg_Path "exefile" -Backup_Path "$Backup_Folder\Backup_HKRoot_EXEFile.reg"
Export_Reg_Config -Reg_Path "Microsoft.PowerShellScript.1" -Backup_Path "$Backup_Folder\Backup_HKRoot_PowerShellScript.reg"
Export_Reg_Config -Reg_Path "VBSFile" -Backup_Path "$Backup_Folder\Backup_HKRoot_VBSFile.reg"
Export_Reg_Config -Reg_Path "Msi.Package" -Backup_Path "$Backup_Folder\Backup_HKRoot_Msi.reg"
Export_Reg_Config -Reg_Path "CompressedFolder" -Backup_Path "$Backup_Folder\Backup_HKRoot_CompressedFolder.reg"
Export_Reg_Config -Reg_Path "WinRAR.ZIP" -Backup_Path "$Backup_Folder\Backup_HKRoot_WinRAR.reg"
Export_Reg_Config -Reg_Path "Directory" -Backup_Path "$Backup_Folder\Backup_HKRoot_Directory.reg"	

Write-Progress -Activity $Progress_Activity -PercentComplete 10

If(-not $NoCheckpoint) 
	{
		Try {
			Checkpoint-Computer -Description "Add Windows Sandbox Context menus" -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop
			Write_Log -Message_Type "SUCCESS" -Message "Creation of restore point `"Add Windows Sandbox Context menus`""
		}
		Catch {
			Write_Log -Message_Type "ERROR" -Message "Creation of restore point `"Add Windows Sandbox Context menus`""
			Write_Log -Message_Type "ERROR" -Message "$($_.Exception.Message)"
		}
	}

Function Add-RegKeys 
{
	param (
		$Reg_Path = "Registry::HKEY_CLASSES_ROOT",
		$Sub_Reg_Path,
		$Type,
		$Entry_Name = $Type,
		$Info_Type = $Type,
		$Key_Label = "Run $Entry_Name in Sandbox"
	)
	
	Try 
		{
			$Base_Registry_Key = "$Reg_Path\$Sub_Reg_Path"
			$Shell_Registry_Key = "$Base_Registry_Key\Shell"
			$Key_Label_Path = "$Shell_Registry_Key\$Key_Label"
			$Command_Path = "$Key_Label_Path\Command"
			$Command_for = "C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -WindowStyle Hidden -NoProfile -ExecutionPolicy Unrestricted -sta -File C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -Type $Type -ScriptPath `"%V`""
			If(-not (Test-Path $Base_Registry_Key)) 
				{
					New-Item $Base_Registry_Key -ErrorAction Stop | Out-Null
				}
			If(-not (Test-Path $Shell_Registry_Key)) 
				{
					New-Item $Shell_Registry_Key -ErrorAction Stop | Out-Null
				}
			If(Test-Path $Key_Label_Path) 
				{
					Write_Log -Message_Type "SUCCESS" -Message "Context menu for $Type has already been added"
					return
				}
				
			New-Item $Key_Label_Path -ErrorAction Stop | Out-Null
			New-Item $Command_Path -ErrorAction Stop | Out-Null
			# Add Sandbox Icons
			New-ItemProperty -Path $Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon -ErrorAction Stop | Out-Null
			# Set the command path
			Set-Item -Path $Command_Path -Value $Command_for -Force -ErrorAction Stop | Out-Null
			Write_Log -Message_Type "SUCCESS" -Message "Context menu for `"$Info_Type`" has been added"
		}
	Catch 
		{
			Write_Log -Message_Type "ERROR" -Message "Context menu for $Type couldn´t be added"
		}
	
}

write-progress -activity $Progress_Activity -percentcomplete 20;

Write_Log -Message_Type "INFO" -Message "Adding context menu"
Write_Log -Message_Type "INFO" -Message "OS version is: $Windows_Version"

If($Add_PS1 -eq $True)
	{
		$PS1_Main_Menu = "Run PS1 in Sandbox"
		
		If($Windows_Version -like "*Windows 10*") 
			{
				Write_Log -Message_Type "INFO" -Message "Running on Windows 10"
				$PS1_Shell_Registry_Key = "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\.ps1\Shell"
				$Main_Menu_Path = "$PS1_Shell_Registry_Key\$PS1_Main_Menu"
				New-Item -Path $PS1_Shell_Registry_Key -Name $PS1_Main_Menu -Force | Out-Null
				New-ItemProperty -Path $Main_Menu_Path -Name "subcommands" -PropertyType String | Out-Null
				New-Item -Path $Main_Menu_Path -Name "Shell" -Force | Out-Null

				Add-RegKeys -Sub_Reg_Path "SystemFileAssociations\.ps1\Shell\Run PS1 in Sandbox" -Type "PS1Basic" -Entry_Name "PS1 as user"
				Add-RegKeys -Sub_Reg_Path "SystemFileAssociations\.ps1\Shell\Run PS1 in Sandbox" -Type "PS1System" -Entry_Name "PS1 as system"
				Add-RegKeys -Sub_Reg_Path "SystemFileAssociations\.ps1\Shell\Run PS1 in Sandbox" -Type "PS1Params" -Entry_Name "PS1 with parameters"
				New-ItemProperty -Path "$Main_Menu_Path" -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null		
			}
		If($Windows_Version -like "*Windows 11*") 
			{
				Write_Log -Message_Type "INFO" -Message "Running on Windows 11"
				
				$Default_PS1_HKCU = "$HKCU_Classes\.ps1"
				If (Test-Path $HKCU_Classes) 
					{
						$rOpenWithProgids_Key = "$Default_PS1_HKCU\rOpenWithProgids"
						If (Test-Path $rOpenWithProgids_Key) 
							{
								$Get_OpenWithProgids_Default_Value = (Get-Item $rOpenWithProgids_Key).Property
								ForEach ($Prop in $Get_OpenWithProgids_Default_Value) 
									{
										$PS1_Shell_Registry_Key = "$HKCU_Classes\$Prop\Shell"
										$Main_Menu_Path = "$PS1_Shell_Registry_Key\$PS1_Main_Menu"
										New-Item -Path $PS1_Shell_Registry_Key -Name $PS1_Main_Menu -Force | Out-Null
										New-ItemProperty -Path $Main_Menu_Path -Name "subcommands" -PropertyType String | Out-Null
										New-Item -Path $Main_Menu_Path -Name "Shell" -Force | Out-Null

										Add-RegKeys -Reg_Path "$PS1_Shell_Registry_Key" -Sub_Reg_Path "$PS1_Main_Menu" -Type "PS1Basic" -Entry_Name "PS1 as user"
										Add-RegKeys -Reg_Path "$PS1_Shell_Registry_Key" -Sub_Reg_Path "$PS1_Main_Menu" -Type "PS1System" -Entry_Name "PS1 as system" 
										Add-RegKeys -Reg_Path "$PS1_Shell_Registry_Key" -Sub_Reg_Path "$PS1_Main_Menu" -Type "PS1Params" -Entry_Name "PS1 with parameters"
									}
							}
						$OpenWithProgids_Key = "$Default_PS1_HKCU\OpenWithProgids"
						If (Test-Path $OpenWithProgids_Key) 
							{
								$Get_OpenWithProgids_Default_Value = (Get-Item $OpenWithProgids_Key).Property
								ForEach ($Prop in $Get_OpenWithProgids_Default_Value) 
									{
										$PS1_Shell_Registry_Key = "$HKCU_Classes\$Prop\Shell"
										$Main_Menu_Path = "$PS1_Shell_Registry_Key\$PS1_Main_Menu"
										New-Item -Path $PS1_Shell_Registry_Key -Name $PS1_Main_Menu -Force | Out-Null
										New-ItemProperty -Path $Main_Menu_Path -Name "subcommands" -PropertyType String | Out-Null

										Add-RegKeys -Reg_Path "$PS1_Shell_Registry_Key" -Sub_Reg_Path "$PS1_Main_Menu" -Type "PS1Basic" -Entry_Name "PS1 as user"
										Add-RegKeys -Reg_Path "$PS1_Shell_Registry_Key" -Sub_Reg_Path "$PS1_Main_Menu" -Type "PS1System" -Entry_Name "PS1 as system" 
										Add-RegKeys -Reg_Path "$PS1_Shell_Registry_Key" -Sub_Reg_Path "$PS1_Main_Menu" -Type "PS1Params" -Entry_Name "PS1 with parameters" 
									
									}
							}
						
						# ADDING CONTEXT MENU DEPENDING OF THE USERCHOICE
						# The userchoice for PS1 is located in: HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.ps1\UserChoice
						$PS1_UserChoice = "$HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.ps1\UserChoice"
						$Get_UserChoice = (Get-ItemProperty $PS1_UserChoice).ProgID

						$HKCR_UserChoice_Key = "Registry::HKEY_CLASSES_ROOT\$Get_UserChoice"
						$PS1_Shell_Registry_Key = "$HKCR_UserChoice_Key\Shell"
						If (Test-Path $PS1_Shell_Registry_Key) 
							{
								$Main_Menu_Path = "$PS1_Shell_Registry_Key\$PS1_Main_Menu"
								New-Item -Path $PS1_Shell_Registry_Key -Name $PS1_Main_Menu -Force | Out-Null
								New-ItemProperty -Path $Main_Menu_Path -Name "subcommands" -PropertyType String | Out-Null

								Add-RegKeys -Reg_Path "$PS1_Shell_Registry_Key" -Sub_Reg_Path "$PS1_Main_Menu" -Type "PS1Basic" -Entry_Name "PS1 as user"
								Add-RegKeys -Reg_Path "$PS1_Shell_Registry_Key" -Sub_Reg_Path "$PS1_Main_Menu" -Type "PS1Params" -Entry_Name "PS1 as system"
								Add-RegKeys -Reg_Path "$PS1_Shell_Registry_Key" -Sub_Reg_Path "$PS1_Main_Menu" -Type "PS1System" -Entry_Name "PS1 with parameters" 
							}
					}
				New-ItemProperty -Path "$Main_Menu_Path" -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null		
			}
	}
Write-Progress -Activity $Progress_Activity -PercentComplete 25

If($Add_PSD1 -eq $True)
	{
		$PSD1_Main_Menu = "Import PowerShell module in Sandbox"
		
		# If($Windows_Version -like "*Windows 10*") 
			# {
				Write_Log -Message_Type "INFO" -Message "Running on Windows 10"
				$PS1_Shell_Registry_Key = "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\.ps1\Shell"
				$Main_Menu_Path = "$PS1_Shell_Registry_Key\$PSD1_Main_Menu"
				New-Item -Path $PS1_Shell_Registry_Key -Name $PSD1_Main_Menu -Force | Out-Null
				New-ItemProperty -Path $Main_Menu_Path -Name "subcommands" -PropertyType String | Out-Null
				New-Item -Path $Main_Menu_Path -Name "Shell" -Force | Out-Null

				Add-RegKeys -Sub_Reg_Path "SystemFileAssociations\.ps1\Shell\Run PS1 in Sandbox" -Type "PS1Basic" -Entry_Name "PS1 as user"
				Add-RegKeys -Sub_Reg_Path "SystemFileAssociations\.ps1\Shell\Run PS1 in Sandbox" -Type "PS1System" -Entry_Name "PS1 as system"
				Add-RegKeys -Sub_Reg_Path "SystemFileAssociations\.ps1\Shell\Run PS1 in Sandbox" -Type "PS1Params" -Entry_Name "PS1 with parameters"
				Add-RegKeys -Sub_Reg_Path "SystemFileAssociations\.ps1\Shell\Run PS1 in Sandbox" -Type "PS1SystemUser" -Entry_Name "PS1 as user in SYSTEM context"

				New-ItemProperty -Path "$Main_Menu_Path" -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null		
			# }
	}
Write-Progress -Activity $Progress_Activity -PercentComplete 25


If($Add_Intunewin -eq $True)
	{
		Add-RegKeys -Sub_Reg_Path ".intunewin" -Type "Intunewin"
	}
Write-Progress -Activity $Progress_Activity -PercentComplete 30
	

If($Add_Reg -eq $True)
	{
		Add-RegKeys -Sub_Reg_Path "regfile" -Type "REG" -Key_Label "Test reg file in Sandbox"
	}
Write-Progress -Activity $Progress_Activity -PercentComplete 35
	
If($Add_ISO -eq $True)
	{
		Add-RegKeys -Sub_Reg_Path "Windows.IsoFile" -Type "ISO" -Key_Label "Extract ISO file in Sandbox"
		Add-RegKeys -Reg_Path "$HKCU_Classes" -Sub_Reg_Path ".iso" -Type "ISO" -Key_Label "Extract ISO file in Sandbox"
	}
Write-Progress -Activity $Progress_Activity -PercentComplete 40

If($Add_PPKG -eq $True)
	{
		Add-RegKeys -Sub_Reg_Path "Microsoft.ProvTool.Provisioning.1" -Type "PPKG"
	}
Write-Progress -Activity $Progress_Activity -PercentComplete 45

If($Add_HTML -eq $True)
	{
		Add-RegKeys -Sub_Reg_Path "MSEdgeHTM" -Type "HTML" -Key_Label "Run this web link in Sandbox"
		Add-RegKeys -Sub_Reg_Path "ChromeHTML" -Type "HTML" -Key_Label "Run this web link in Sandbox"
		Add-RegKeys -Sub_Reg_Path "IE.AssocFile.HTM" -Type "HTML" -Key_Label "Run this web link in Sandbox"
		Add-RegKeys -Sub_Reg_Path "IE.AssocFile.URL" -Type "HTML" -Key_Label "Run this URL in Sandbox"
	}
Write-Progress -Activity $Progress_Activity -PercentComplete 50

If($Add_MultipleApp -eq $True)
	{
		Add-RegKeys -Sub_Reg_Path ".sdbapp" -Type "SDBApp" -Entry_Name "application bundle"
	}
Write-Progress -Activity $Progress_Activity -PercentComplete 55																

If($Add_VBS -eq $True)
	{
		$VBS_Main_Menu = "Run VBS in Sandbox"

		$VBS_Shell_Registry_Key = "Registry::HKEY_CLASSES_ROOT\VBSFile\Shell"
		$Main_Menu_Path = "$VBS_Shell_Registry_Key\$VBS_Main_Menu"
		New-Item -Path $VBS_Shell_Registry_Key -Name $VBS_Main_Menu -Force | Out-Null
		New-ItemProperty -Path $Main_Menu_Path -Name "subcommands" -PropertyType String | Out-Null
		New-Item -Path $Main_Menu_Path -Name "Shell" -Force | Out-Null

		Add-RegKeys -Sub_Reg_Path "VBSFile\Shell\$VBS_Main_Menu" -Type "VBSBasic" -Entry_Name "VBS"
		Add-RegKeys -Sub_Reg_Path "VBSFile\Shell\$VBS_Main_Menu" -Type "VBSParams" -Entry_Name "VBS with parameters"
		
		New-ItemProperty -Path "$Main_Menu_Path" -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null				
	}
Write-Progress -Activity $Progress_Activity -PercentComplete 60


If($Add_EXE -eq $True)
	{
		Add-RegKeys -Sub_Reg_Path "exefile" -Type "EXE"
	}
Write-Progress -Activity $Progress_Activity -PercentComplete 65

If($Add_MSI -eq $True)
	{
		Add-RegKeys -Sub_Reg_Path "Msi.Package" -Type "MSI"
	}
Write-Progress -Activity $Progress_Activity -PercentComplete 70

If($Add_ZIP -eq $True)
	{
		#Run on ZIP
		Add-RegKeys -Sub_Reg_Path "CompressedFolder" -Type "ZIP" -Key_Label "Extract ZIP in Sandbox"

		# Run on ZIP if WinRAR is installed
		If (Test-Path "Registry::HKEY_CLASSES_ROOT\WinRAR.ZIP") {
			Add-RegKeys -Sub_Reg_Path "WinRAR.ZIP" -Type "ZIP" -Key_Label "Extract ZIP in Sandbox"
		}

		# Run on 7z
		If (Test-Path "Registry::HKEY_CLASSES_ROOT\Applications\7zFM.exe") {
			Add-RegKeys -Sub_Reg_Path "Applications\7zFM.exe" -Type "ZIP" -Info_Type "7z" -Entry_Name "ZIP" -Key_Label "Extract 7z file in Sandbox"
		}
	}
Write-Progress -Activity $Progress_Activity -PercentComplete 75
		
If($Add_MSIX -eq $True)
	{
		$MSIX_Shell_Registry_Key = "Registry::HKEY_CLASSES_ROOT\.msix\OpenWithProgids"
		If (Test-Path $MSIX_Shell_Registry_Key) {
			$Get_Default_Value = (Get-Item $MSIX_Shell_Registry_Key).Property
			Add-RegKeys -Sub_Reg_Path "$Get_Default_Value" -Type "MSIX"
		}
		If (Test-Path $HKCU_Classes) {
			$Default_MSIX_HKCU = "$HKCU_Classes\.msix"
			$Get_Default_Value = (Get-Item "$Default_MSIX_HKCU\OpenWithProgids").Property
			Add-RegKeys -Reg_Path $HKCU_Classes -Sub_Reg_Path "$Get_Default_Value" -Type "MSIX"
		} 
	}
Write-Progress -Activity $Progress_Activity -PercentComplete 80
	
If($Add_Folder -eq $True)
	{
		Add-RegKeys -Sub_Reg_Path "Directory\Background" -Type "Folder_Inside" -Entry_Name "this folder" -Key_Label "Share this folder in a Sandbox"
		Add-RegKeys -Sub_Reg_Path "Directory" -Type "Folder_On" -Entry_Name "this folder" -Key_Label "Share this folder in a Sandbox"
	}
Write-Progress -Activity $Progress_Activity -PercentComplete 85

If($Add_CMD -eq $True)
	{
		Add-RegKeys -Sub_Reg_Path "cmdfile" -Type "CMD"
		Add-RegKeys -Sub_Reg_Path "batfile" -Type "BAT"
	}
Write-Progress -Activity $Progress_Activity -PercentComplete 90


If($Add_PDF -eq $True)
	{

		If(-not (Test-Path "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\.pdf")) 
			{
				New-Item "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\.pdf" -ErrorAction Stop | Out-Null
			}
		Add-RegKeys -Sub_Reg_Path "SystemFileAssociations\.pdf" -Type "PDF" -Key_Label "Open PDF in Sandbox"
	}
Write-Progress -Activity $Progress_Activity -PercentComplete 100

copy-item $Log_File $Destination_folder -Force																
