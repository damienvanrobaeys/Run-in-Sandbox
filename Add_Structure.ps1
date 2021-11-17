Param
(
	[Switch]$NoSilent
)

#***************************************************************************************************************
# Author: Damien VAN ROBAEYS
# Website: http://www.systanddeploy.com
# Twitter: https://twitter.com/syst_and_deploy
#***************************************************************************************************************
$TEMP_Folder = $env:temp
$Log_File = "$TEMP_Folder\RunInSandbox_Install.log"
$Current_Folder = split-path $MyInvocation.MyCommand.Path

If(test-path $Log_File){remove-item $Log_File}
new-item $Log_File -type file -force | out-null
Function Write_Log
	{
		param(
		$Message_Type,	
		$Message
		)
		
		$MyDate = "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)		
		# Add-Content $Log_File  "$MyDate - $Message_Type : $Message"	
		write-host "$MyDate - $Message_Type : $Message"	
	}
	
Function Export_Reg_Config
	{
		param(
		$Reg_Path,
		$Backup_Path
		)
		
		Write_Log -Message_Type "INFO" -Message "Exporting registry path: $Reg_Path"		
				
		If(test-path "HKCR_SD:\$Reg_Path")
			{
				Try
					{
						reg export "HKEY_CLASSES_ROOT\$Reg_Path" $Backup_Path /y | out-null
						Write_Log -Message_Type "SUCCESS" -Message "$Reg_Path has been exported"									
					}
				Catch
					{
						Write_Log -Message_Type "ERROR" -Message "$Reg_Path has not been exported"									
					}			
			}
		Else
			{
				Write_Log -Message_Type "INFO" -Message "Can not find registry path: HKCR_SD:\$Reg_Path"															
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
Else
	{	
		Write_Log -Message_Type "INFO" -Message "The script has been launched with admin rights"

		$Is_Sandbox_Installed = (Get-WindowsOptionalFeature -online | where {$_.featurename -eq "Containers-DisposableClientVM"}).state
		If($Is_Sandbox_Installed -eq "Disabled")
			{
				Write_Log -Message_Type "INFO" -Message "The feature Windows Sandbox is not installed !!!"	
				[System.Windows.Forms.MessageBox]::Show("The feature Windows Sandbox is not installed !!!")		
			}
		Else
			{
				$Current_Folder = split-path $MyInvocation.MyCommand.Path
				$Sources = $Current_Folder + "\" + "Sources\*"
				If(test-path $Sources)
					{
						Write_Log -Message_Type "INFO" -Message "The sources folder exists"		
						
						Add-content $log_file ""		

						$Progress_Activity = "Enabling Run in Sandbox context menus"
						write-progress -activity $Progress_Activity -percentcomplete 1;
								
						$Check_Sources_Files_Count = (get-childitem "$Current_Folder\Sources\Run_in_Sandbox" -recurse).count
						# If($Check_Sources_Files_Count -eq 36)	
						If($Check_Sources_Files_Count -eq 38)																			
							{	
								$Sources_Copied = $False
								$ProgData = $env:ProgramData
								$Destination_folder = "$ProgData\Run_in_Sandbox"
								Try
									{
										copy-item $Sources $ProgData -force -recurse | out-null						
										Write_Log -Message_Type "SUCCESS" -Message "Sources have been copied in $ProgData\Run_in_Sandbox"	
										$Sources_Copied = $True								
									}
								Catch
									{
										Write_Log -Message_Type "ERROR" -Message "Sources have not been copied in $ProgData\Run_in_Sandbox"								
										Break
									}

								If($Sources_Copied -eq $True)
									{
										$Sources_Unblocked = $False							
										Try
											{
												Get-Childitem -Recurse $Destination_folder | Unblock-file
												Write_Log -Message_Type "SUCCESS" -Message "Sources files have been unblocked"																	
												$Sources_Unblocked = $True
											}
										Catch	
											{
												Write_Log -Message_Type "SUCCESS" -Message "Sources files have not been unblocked"	
												Break
											}
											
										If($Sources_Unblocked -eq $True)
											{
												If($NoSilent)
													{
														cd "$Current_Folder\Sources\Run_in_Sandbox"
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
																																							
												If(test-path "$ProgData\Run_in_Sandbox\RunInSandbox.ps1")
													{					
														$Backup_Folder = "$Destination_folder\Registry_Backup"
														new-item $Backup_Folder -Type Directory -Force | out-null
														write-progress -activity $Progress_Activity  -percentcomplete 5;
														
														$List_Drive = get-psdrive | where {$_.Name -eq "HKCR_SD"}
														If($List_Drive -ne $null){Remove-PSDrive $List_Drive}
														
														Write_Log -Message_Type "INFO" -Message "Mapping registry HKCR"													
														
														$HKCR_Mapped = $False
														Try
															{
																New-PSDrive -PSProvider registry -Root HKEY_CLASSES_ROOT -Name HKCR_SD | out-null
																Write_Log -Message_Type "SUCCESS" -Message "Mapping registry HKCR"		
																$HKCR_Mapped = $True												
															}
														Catch
															{
																Write_Log -Message_Type "ERROR" -Message "Mapping registry HKCR"
																Break
															}

														If($HKCR_Mapped -eq $True)
															{
																write-progress -activity $Progress_Activity  -percentcomplete 10;

																Export_Reg_Config -Reg_Path "exefile" -Backup_Path "$Backup_Folder\Backup_HKRoot_EXEFile.reg"
																Export_Reg_Config -Reg_Path "Microsoft.PowerShellScript.1" -Backup_Path "$Backup_Folder\Backup_HKRoot_PowerShellScript.reg"
																Export_Reg_Config -Reg_Path "VBSFile" -Backup_Path "$Backup_Folder\Backup_HKRoot_VBSFile.reg"
																Export_Reg_Config -Reg_Path "Msi.Package" -Backup_Path "$Backup_Folder\Backup_HKRoot_Msi.reg"
																Export_Reg_Config -Reg_Path "CompressedFolder" -Backup_Path "$Backup_Folder\Backup_HKRoot_CompressedFolder.reg"
																Export_Reg_Config -Reg_Path "WinRAR.ZIP" -Backup_Path "$Backup_Folder\Backup_HKRoot_WinRAR.reg"
																Export_Reg_Config -Reg_Path "Directory" -Backup_Path "$Backup_Folder\Backup_HKRoot_Directory.reg"						

																write-progress -activity $Progress_Activity  -percentcomplete 15;	
																
																Add-content $log_file ""		
																
																Write_Log -Message_Type "INFO" -Message "Creating a restore point"		
																Checkpoint-Computer -Description "Add Windows Sandbox Context menus" -RestorePointType "MODIFY_SETTINGS" -ea silentlycontinue -ev ErrorRestore
																If($ErrorRestore -ne $null)
																	{
																		Write_Log -Message_Type "SUCCESS" -Message "Creation of restore point 'Add Windows Sandbox Context menus'"																
																	}
																Else
																	{
																		Write_Log -Message_Type "ERROR" -Message "Creation of restore point 'Add Windows Sandbox Context menus'"	
																		Write_Log -Message_Type "ERROR" -Message "$ErrorRestore"																														
																	}
																
																Add-content $log_file ""		

																# $Main_Language = $Get_XML_Content.Configuration.Main_Language
																# If($Main_Language -ne $null)
																	# {
																		# $Get_lang_to_install = $Get_XML_Content.Configuration.Main_Language	
																	# }
																# Else
																	# {
																		# $Get_lang_to_install = (Get-Culture).name
																	# }
																# $Language_File = (Get-Childitem "$Current_Folder\Sources\Run_in_Sandbox\Languages_XML" | Where {$_.Basename -like "*$Get_lang_to_install*"}).fullname
																# $Get_Language_File_Content = ([xml](get-content $Language_File)).Configuration

																write-progress -activity $Progress_Activity -percentcomplete 20;

																If($Add_PS1 -eq $True)
																	{
																		$PS1_Main_Menu = "Run PS1 in Sandbox"
																		$PS1_SubMenu_RunAsUser = "Run PS1 as user"
																		$PS1_SubMenu_RunAsSystem = "Run PS1 as system"
																		$PS1_SubMenu_RunwithParams = "Run PS1 with parameters"
																		
																		$Command_For_Basic_PS1 = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type PS1Basic -LiteralPath "%V" -ScriptPath "%V"' 
																		$Command_For_Params_PS1 = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type PS1Params -LiteralPath "%V" -ScriptPath "%V"' 
																		$Command_For_System_PS1 = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type PS1System -LiteralPath "%V" -ScriptPath "%V"' 
																																				
																		$Windows_Version = (Get-WmiObject -class Win32_OperatingSystem).Caption
																		If($Windows_Version -like "*Windows 10*")
																			{
																				# RUN ON PS1
																				$PS1_Shell_Registry_Key = "HKCR_SD:\Microsoft.PowerShellScript.1\Shell"
																				$Main_Menu_Path = "$PS1_Shell_Registry_Key\$PS1_Main_Menu"
																				New-Item -Path $PS1_Shell_Registry_Key -Name $PS1_Main_Menu -force | out-null
																				New-ItemProperty -Path $Main_Menu_Path -Name "subcommands" -PropertyType String | out-null

																				New-Item -Path $Main_Menu_Path -Name "Shell" -force | out-null
																				$Main_Menu_Shell_Path = "$Main_Menu_Path\Shell"

																				# New-Item -Path $Main_Menu_Shell_Path -Name "Shell" -force | out-null
																				New-Item -Path $Main_Menu_Shell_Path -Name $PS1_SubMenu_RunAsUser -force | out-null
																				New-Item -Path $Main_Menu_Shell_Path -Name $PS1_SubMenu_RunAsSystem -force | out-null
																				New-Item -Path $Main_Menu_Shell_Path -Name $PS1_SubMenu_RunwithParams -force | out-null

																				New-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsUser" -Name "Command" -force | out-null
																				New-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsSystem" -Name "Command" -force | out-null
																				New-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunwithParams" -Name "Command" -force | out-null

																				Set-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsUser\command" -Value $Command_For_Basic_PS1 -force | out-null
																				Set-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunwithParams\command" -Value $Command_For_Params_PS1 -force | out-null
																				Set-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsSystem\command" -Value $Command_For_System_PS1 -force | out-null

																				New-ItemProperty -Path "$Main_Menu_Path" -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null
																				New-ItemProperty -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunwithParams" -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null
																				New-ItemProperty -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsUser" -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null
																				New-ItemProperty -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsSystem" -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null																			
																			}
																		ElseIf($Windows_Version -like "*Windows 11*")
																			{
																				$Current_User_SID = (Get-ChildItem Registry::\HKEY_USERS | Where-Object { Test-Path "$($_.pspath)\Volatile Environment" } | ForEach-Object { (Get-ItemProperty "$($_.pspath)\Volatile Environment")}).PSParentPath.split("\")[-1]																			# RUN ON ISO
																				$HKCU_Classes = "Registry::HKEY_USERS\$Current_User_SID" + "_Classes"
																				If(test-path $HKCU_Classes)
																				{
																					$Default_PS1_HKCU = "$HKCU_Classes\.ps1"
																					$rOpenWithProgids_Key = "$Default_PS1_HKCU\rOpenWithProgids"
																					If(test-path $rOpenWithProgids_Key)
																						{
																							$Get_rOpenWithProgids_Default_Value = (Get-Item $rOpenWithProgids_Key).Property
																							ForEach($Prop in $Get_OpenWithProgids_Default_Value)
																								{
																									$Default_HKCU_PS1_Shell_Registry_Key = "$HKCU_Classes\$Prop\Shell"
																									If(test-path $Default_HKCU_PS1_Shell_Registry_Key)
																										{
																											$Main_Menu_Path = "$Default_HKCU_PS1_Shell_Registry_Key\$PS1_Main_Menu"
																											New-Item -Path $Default_HKCU_PS1_Shell_Registry_Key -Name $PS1_Main_Menu -force | out-null
																											New-ItemProperty -Path $Main_Menu_Path -Name "subcommands" -PropertyType String | out-null

																											New-Item -Path $Main_Menu_Path -Name "Shell" -force | out-null
																											$Main_Menu_Shell_Path = "$Main_Menu_Path\Shell"

																											New-Item -Path $Main_Menu_Shell_Path -Name $PS1_SubMenu_RunAsUser -force | out-null
																											New-Item -Path $Main_Menu_Shell_Path -Name $PS1_SubMenu_RunAsSystem -force | out-null
																											New-Item -Path $Main_Menu_Shell_Path -Name $PS1_SubMenu_RunwithParams -force | out-null

																											New-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsUser" -Name "Command" -force | out-null
																											New-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsSystem" -Name "Command" -force | out-null
																											New-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunwithParams" -Name "Command" -force | out-null

																											Set-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsUser\command" -Value $Command_For_Basic_PS1 -force | out-null
																											Set-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunwithParams\command" -Value $Command_For_Params_PS1 -force | out-null
																											Set-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsSystem\command" -Value $Command_For_System_PS1 -force | out-null

																											# Add Sandbox Icon
																											New-ItemProperty -Path "$Main_Menu_Path" -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null
																											New-ItemProperty -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunwithParams" -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null
																											New-ItemProperty -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsUser" -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null
																											New-ItemProperty -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsSystem" -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null	
																										}																									
																								}																						
																						}
	
																					$OpenWithProgids_Key = "$Default_PS1_HKCU\OpenWithProgids"
																					If(test-path $OpenWithProgids_Key)
																						{
																							$Get_OpenWithProgids_Default_Value = (Get-Item $OpenWithProgids_Key).Property
																							ForEach($Prop in $Get_OpenWithProgids_Default_Value)
																								{
																									$Default_HKCU_PS1_Shell_Registry_Key = "$HKCU_Classes\$Prop\Shell"
																									If(test-path $Default_HKCU_PS1_Shell_Registry_Key)
																										{
																											$Main_Menu_Path = "$Default_HKCU_PS1_Shell_Registry_Key\$PS1_Main_Menu"
																											New-Item -Path $Default_HKCU_PS1_Shell_Registry_Key -Name $PS1_Main_Menu -force | out-null
																											New-ItemProperty -Path $Main_Menu_Path -Name "subcommands" -PropertyType String | out-null

																											New-Item -Path $Main_Menu_Path -Name "Shell" -force | out-null
																											$Main_Menu_Shell_Path = "$Main_Menu_Path\Shell"

																											New-Item -Path $Main_Menu_Shell_Path -Name $PS1_SubMenu_RunAsUser -force | out-null
																											New-Item -Path $Main_Menu_Shell_Path -Name $PS1_SubMenu_RunAsSystem -force | out-null
																											New-Item -Path $Main_Menu_Shell_Path -Name $PS1_SubMenu_RunwithParams -force | out-null

																											New-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsUser" -Name "Command" -force | out-null
																											New-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsSystem" -Name "Command" -force | out-null
																											New-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunwithParams" -Name "Command" -force | out-null

																											Set-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsUser\command" -Value $Command_For_Basic_PS1 -force | out-null
																											Set-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunwithParams\command" -Value $Command_For_Params_PS1 -force | out-null
																											Set-Item -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsSystem\command" -Value $Command_For_System_PS1 -force | out-null

																											# Add Sandbox Icon
																											New-ItemProperty -Path "$Main_Menu_Path" -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null
																											New-ItemProperty -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunwithParams" -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null
																											New-ItemProperty -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsUser" -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null
																											New-ItemProperty -Path "$Main_Menu_Shell_Path\$PS1_SubMenu_RunAsSystem" -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null	
																										}																									
																								}																						
																						}																						
																				}
																			}																			
																		Write_Log -Message_Type "INFO" -Message "Context menus for PS1 have been added"																		
																	}
																	
																write-progress -activity $Progress_Activity  -percentcomplete 25;

																If($Add_Intunewin -eq $True)
																	{
																		# RUN ON INTUNEWIN
																		$Intunewin_Shell_Registry_Key = "HKCR_SD:\.intunewin"
																		$Intunewin_Key_Label = "Test intunewin in Sandbox"
																		$Intunewin_Key_Label_Path = "$Intunewin_Shell_Registry_Key\Shell\$Intunewin_Key_Label"
																		$Intunewin_Command_Path = "$Intunewin_Key_Label_Path\Command"
																		$Command_for_Intunewin = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -windowstyle hidden -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type Intunewin -LiteralPath "%V" -ScriptPath "%V"'
																		If(!(test-path $Intunewin_Shell_Registry_Key))
																		{
																			new-item $Intunewin_Shell_Registry_Key -force | out-null
																			new-item "$Intunewin_Shell_Registry_Key\Shell" -force | out-null
																		}
																		new-item $Intunewin_Key_Label_Path -force | out-null
																		new-item $Intunewin_Command_Path -force | out-null	
																		# Set the command path
																		Set-Item -Path $Intunewin_Command_Path -Value $Command_for_Intunewin -force | out-null	
																		# Add Sandbox Icons
																		New-ItemProperty -Path $Intunewin_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon -force | out-null			
																		Write_Log -Message_Type "INFO" -Message "Context menus for PS1 have been added"																									
																																			
																	}
																	
																write-progress -activity $Progress_Activity  -percentcomplete 30;		
																
																If($Add_Reg -eq $True)
																		{
																			# RUN ON REG
																			$Reg_Shell_Registry_Key = "HKCR_SD:\regfile\Shell"
																			$Reg_Key_Label = "Test reg file in Sandbox"
																			$Reg_Key_Label_Path = "$Reg_Shell_Registry_Key\$Reg_Key_Label"
																			$Reg_Command_Path = "$Reg_Key_Label_Path\Command"
																			$Command_for_Reg = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -windowstyle hidden -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type REG -LiteralPath "%V" -ScriptPath "%V"'
																			If(test-path $Reg_Shell_Registry_Key)
																			{
																				new-item $Reg_Key_Label_Path | out-null
																				new-item $Reg_Command_Path | out-null	
																				# Set the command path
																				Set-Item -Path $Reg_Command_Path -Value $Command_for_Reg -force | out-null	
																				# Add Sandbox Icons
																				New-ItemProperty -Path $Reg_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null			
																				Write_Log -Message_Type "INFO" -Message "Context menus for REG have been added"																									
																			}																	
																		}

																write-progress -activity $Progress_Activity  -percentcomplete 35;																		
																	
																If($Add_ISO -eq $True)
																		{
																			$Current_User_SID = (Get-ChildItem Registry::\HKEY_USERS | Where-Object { Test-Path "$($_.pspath)\Volatile Environment" } | ForEach-Object { (Get-ItemProperty "$($_.pspath)\Volatile Environment")}).PSParentPath.split("\")[-1]																			# RUN ON ISO

																			$ISO_Key_Label = "Extract ISO file in Sandbox"
																			$Command_for_ISO = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -windowstyle hidden -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type ISO -LiteralPath "%V" -ScriptPath "%V"'																			
																			
																			# Modify value from HKCR
																			$ISO_Shell_Registry_Key = "HKCR_SD:\Windows.IsoFile\Shell"
																			$ISO_Key_Label_Path = "$ISO_Shell_Registry_Key\$ISO_Key_Label"
																			$ISO_Command_Path = "$ISO_Key_Label_Path\Command"
																			If(test-path $ISO_Shell_Registry_Key)
																			{
																				new-item $ISO_Key_Label_Path | out-null
																				new-item $ISO_Command_Path | out-null	
																				# Set the command path
																				Set-Item -Path $ISO_Command_Path -Value $Command_for_ISO -force | out-null	
																				# Add Sandbox Icons
																				New-ItemProperty -Path $ISO_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null			
																				Write_Log -Message_Type "INFO" -Message "Context menu for ISO has been added"																															
																			}			
																			
																			Write_Log -Message_Type "INFO" -Message "Checking content of HKCR\.ISO"																																																		
																			$ISO_Key = "HKCR_SD:\.ISO"
																			If(test-path $ISO_Key)
																				{
																					Write_Log -Message_Type "INFO" -Message "The key HKCR\.ISO exists"	
																					$Get_ISO_Keys = Get-Item $ISO_Key
																					ForEach($Key in $Get_ISO_Keys)
																					{
																						$Get_Properties = $Key.Property
																						Write_Log -Message_Type "INFO" -Message "Following subkeys found: $Get_Properties"	
																						foreach($Property in $Get_Properties)
																							{
																								$Prop = (Get-ItemProperty $ISO_Key)."$Property"
																								Write_Log -Message_Type "INFO" -Message "Following property found: $Prop"	
																								$ISO_Property_Key = "HKCR_SD:\$Prop"
																								Write_Log -Message_Type "INFO" -Message "Reg path to test: $ISO_Property_Key"	
																								If(test-path $ISO_Property_Key)
																									{
																										Write_Log -Message_Type "INFO" -Message "The following reg path exists: $ISO_Property_Key"
																										$ISO_Property_Shell = "$ISO_Property_Key\Shell"
																										If(!(test-path $ISO_Property_Shell))
																											{
																												new-item $ISO_Property_Shell | out-null
																											}	
																										$ISO_Key_Label_Path = "$ISO_Property_Shell\$ISO_Key_Label"
																										$ISO_Command_Path = "$ISO_Key_Label_Path\Command"
																										new-item $ISO_Key_Label_Path | out-null
																										new-item $ISO_Command_Path | out-null	
																										# Set the command path
																										Set-Item -Path $ISO_Command_Path -Value $Command_for_ISO -force | out-null	
																										# Add Sandbox Icons
																										New-ItemProperty -Path $ISO_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null			
																										Write_Log -Message_Type "INFO" -Message "Creating following context menu for ISO under: $ISO_Key_Label_Path"																																																					
																									}
																									Else
																									{
																										Write_Log -Message_Type "INFO" -Message "The following reg path does not exist: $ISO_Property_Shell"
																									}
																							}																					
																					}																				
																				}																			

																			

																			# Modify value from HKCU if 7zip exists
																			$HKCU_Classes = "Registry::HKEY_USERS\$Current_User_SID" + "_Classes"
																			If(test-path $HKCU_Classes)
																			{
																				$Default_ISO_HKCU = "$HKCU_Classes\.iso"
																				If(test-path $Default_ISO_HKCU)
																					{
																						$Get_Default_Value = (Get-ItemProperty $Default_ISO_HKCU)."(default)"
																						$Default_HKCU_ISO_Shell_Registry_Key = "$HKCU_Classes\$Get_Default_Value\Shell"
																						If(test-path $Default_HKCU_ISO_Shell_Registry_Key)
																							{
																								$ISO_HKCU_Key_Label_Path = "$Default_HKCU_ISO_Shell_Registry_Key\$ISO_Key_Label"
																								$HKCU_ISO_Command_Path = "$ISO_HKCU_Key_Label_Path\Command"																					
																								new-item $ISO_HKCU_Key_Label_Path | out-null
																								new-item $HKCU_ISO_Command_Path | out-null	
																								# Set the command path
																								Set-Item -Path $HKCU_ISO_Command_Path -Value $Command_for_ISO -force | out-null	
																								# Add Sandbox Icons
																								New-ItemProperty -Path $ISO_HKCU_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null			
																								Write_Log -Message_Type "INFO" -Message "Context menu for ISO has been added"																																													
																							}												
																					}
																																																						
																			}																				
																		}

																write-progress -activity $Progress_Activity  -percentcomplete 40;																		


																If($Add_PPKG -eq $True)
																		{
																			# RUN ON PPKG
																			$PPKG_Shell_Registry_Key = "HKCR_SD:\Microsoft.ProvTool.Provisioning.1\Shell"
																			$PPKG_Key_Label = "Run PPKG file in Sandbox"
																			$PPKG_Key_Label_Path = "$PPKG_Shell_Registry_Key\$PPKG_Key_Label"
																			$PPKG_Command_Path = "$PPKG_Key_Label_Path\Command"
																			$Command_for_PPKG = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -windowstyle hidden -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type PPKG -LiteralPath "%V" -ScriptPath "%V"'
																			If(test-path $PPKG_Shell_Registry_Key)
																			{
																				new-item $PPKG_Key_Label_Path | out-null
																				new-item $PPKG_Command_Path | out-null	
																				# Set the command path
																				Set-Item -Path $PPKG_Command_Path -Value $Command_for_PPKG -force | out-null	
																				# Add Sandbox Icons
																				New-ItemProperty -Path $PPKG_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null			
																				Write_Log -Message_Type "INFO" -Message "Context menu for PPKG has been added"																									
																			}																	
																		}

																write-progress -activity $Progress_Activity  -percentcomplete 40;		


																If($Add_HTML -eq $True)
																		{
																			$HTML_Key_Label = "Run this web link in Sandbox"
																			
																			# RUN ON HTML for Edge
																			$HTML_Edge_Shell_Registry_Key = "HKCR_SD:\MSEdgeHTM\Shell"
																			If(test-path $HTML_Edge_Shell_Registry_Key)
																				{
																					$HTML_Edge_Key_Label_Path = "$HTML_Edge_Shell_Registry_Key\$HTML_Key_Label"
																					$HTML_Edge_Command_Path = "$HTML_Edge_Key_Label_Path\Command"
																					$Command_for_HTML = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -windowstyle hidden -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type HTML -LiteralPath "%V" -ScriptPath "%V"'
																					new-item $HTML_Edge_Key_Label_Path | out-null
																					new-item $HTML_Edge_Command_Path | out-null	
																					# Set the command path
																					Set-Item -Path $HTML_Edge_Command_Path -Value $Command_for_HTML -force | out-null	
																					# Add Sandbox Icons
																					New-ItemProperty -Path $HTML_Edge_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null			
																					Write_Log -Message_Type "INFO" -Message "Context menu for HTML has been added"																									
																				}
																				
																			# RUN ON HTML for Chrome
																			$HTML_Chrome_Shell_Registry_Key = "HKCR_SD:\ChromeHTML\Shell"
																			If(test-path $HTML_Chrome_Shell_Registry_Key)
																				{
																					$HTML_Chrome_Key_Label_Path = "$HTML_Chrome_Shell_Registry_Key\$HTML_Key_Label"
																					$HTML_Chrome_Command_Path = "$HTML_Chrome_Key_Label_Path\Command"
																					$Command_for_HTML = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -windowstyle hidden -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type HTML -LiteralPath "%V" -ScriptPath "%V"'

																					new-item $HTML_Chrome_Key_Label_Path | out-null
																					new-item $HTML_Chrome_Command_Path | out-null	
																					# Set the command path
																					Set-Item -Path $HTML_Chrome_Command_Path -Value $Command_for_HTML -force | out-null	
																					# Add Sandbox Icons
																					New-ItemProperty -Path $HTML_Chrome_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null			
																					Write_Log -Message_Type "INFO" -Message "Context menu for HTML has been added"																									
																				}

																			# RUN ON HTML for IE
																			$HTML_IE_Shell_Registry_Key = "HKCR_SD:\IE.AssocFile.HTM\Shell"
																			If(test-path $HTML_IE_Shell_Registry_Key)
																				{
																					$HTML_IE_Key_Label_Path = "$HTML_IE_Shell_Registry_Key\$HTML_Key_Label"
																					$HTML_IE_Command_Path = "$HTML_IE_Key_Label_Path\Command"
																					$Command_for_HTML = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -windowstyle hidden -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type HTML -LiteralPath "%V" -ScriptPath "%V"'
																					new-item $HTML_IE_Key_Label_Path | out-null
																					new-item $HTML_IE_Command_Path | out-null	
																					# Set the command path
																					Set-Item -Path $HTML_IE_Command_Path -Value $Command_for_HTML -force | out-null	
																					# Add Sandbox Icons
																					New-ItemProperty -Path $HTML_IE_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null			
																					Write_Log -Message_Type "INFO" -Message "Context menu for HTML has been added"																									
																				}

																			$URL_Shell_Registry_Key = "HKCR_SD:\IE.AssocFile.URL\Shell"
																			If(test-path $URL_Shell_Registry_Key)
																				{
																					$URL_Key_Label_Path = "$URL_Shell_Registry_Key\Run this URL in Sandbox"
																					$URL_Command_Path = "$URL_Key_Label_Path\Command"
																					$Command_for_URL = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -windowstyle hidden -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type URL -LiteralPath "%V" -ScriptPath "%V"'
																					new-item $URL_Key_Label_Path | out-null
																					new-item $URL_Command_Path | out-null	
																					# Set the command path
																					Set-Item -Path $URL_Command_Path -Value $Command_for_URL -force | out-null	
																					# Add Sandbox Icons
																					New-ItemProperty -Path $URL_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null			
																					Write_Log -Message_Type "INFO" -Message "Context menu for URL has been added"																									
																				}
																		}

																write-progress -activity $Progress_Activity  -percentcomplete 45;		
		
																If($Add_MultipleApp -eq $True)
																	{
																		# RUN ON bundle app
																		$MultipleApps_Shell_Registry_Key = "HKCR_SD:\.sdbapp"
																		$MultipleApps_Key_Label = "Test application bundle in Sandbox"
																		$MultipleApps_Key_Label_Path = "$MultipleApps_Shell_Registry_Key\Shell\$MultipleApps_Key_Label"
																		$MultipleApps_Command_Path = "$MultipleApps_Key_Label_Path\Command"
																		$Command_for_MultipleApps = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -windowstyle hidden -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type SDBApp -LiteralPath "%V" -ScriptPath "%V"'
																		If(!(test-path $MultipleApps_Shell_Registry_Key))
																		{
																			new-item $MultipleApps_Shell_Registry_Key | out-null
																			new-item "$MultipleApps_Shell_Registry_Key\Shell" | out-null
																			new-item $MultipleApps_Key_Label_Path | out-null
																			new-item $MultipleApps_Command_Path | out-null	
																			# Set the command path
																			Set-Item -Path $MultipleApps_Command_Path -Value $Command_for_MultipleApps -force | out-null	
																			# Add Sandbox Icons
																			New-ItemProperty -Path $MultipleApps_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null			
																			Write_Log -Message_Type "INFO" -Message "Context menu for PS1 has been added"																									
																		}																	
																	}																	
																
																write-progress -activity $Progress_Activity  -percentcomplete 45;

																If($Add_VBS -eq $True)
																	{
																		# RUN ON VBS
																		$VBS_Shell_Registry_Key = "HKCR_SD:\VBSFile\Shell"																		
																		$VBS_Basic_Run = "Run VBS in Sandbox"
																		$VBS_Parameter_Run = "Run VBS in Sandbox with parameters"					
																		
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
																		Write_Log -Message_Type "INFO" -Message "Context menus for VBS have been added"																		
																	}

																write-progress -activity $Progress_Activity  -percentcomplete 50;

																If($Add_EXE -eq $True)
																	{
																		# RUN ON EXE
																		$EXE_Shell_Registry_Key = "HKCR_SD:\exefile\Shell"
																		$EXE_Basic_Run = "Run EXE in Sandbox"					
																		$ContextMenu_Basic_EXE = "$EXE_Shell_Registry_Key\$EXE_Basic_Run"				

																		New-Item -Path $EXE_Shell_Registry_Key -Name $EXE_Basic_Run -force | out-null
																		New-Item -Path $ContextMenu_Basic_EXE -Name "Command" -force | out-null
																		# Add Sandbox Icons
																		New-ItemProperty -Path $ContextMenu_Basic_EXE -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null
																		$Command_For_EXE = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type EXE -LiteralPath "%V" -ScriptPath "%V"' 
																		# Set the command path
																		Set-Item -Path "$ContextMenu_Basic_EXE\command" -Value $Command_For_EXE -force | out-null
																		Write_Log -Message_Type "INFO" -Message "Context menus for EXE have been added"																		
																	}

																write-progress -activity $Progress_Activity  -percentcomplete 50;

																If($Add_MSI -eq $True)
																	{
																		# RUN ON MSI
																		$MSI_Shell_Registry_Key = "HKCR_SD:\Msi.Package\Shell"
																		$MSI_Basic_Run = "Run MSI in Sandbox"					
																		$ContextMenu_Basic_MSI = "$MSI_Shell_Registry_Key\$MSI_Basic_Run"				

																		New-Item -Path $MSI_Shell_Registry_Key -Name $MSI_Basic_Run -force | out-null
																		New-Item -Path $ContextMenu_Basic_MSI -Name "Command" -force | out-null
																		# Add Sandbox Icons
																		New-ItemProperty -Path $ContextMenu_Basic_MSI -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null
																		$Command_For_MSI = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type MSI -LiteralPath "%V" -ScriptPath "%V"' 
																		# Set the command path
																		Set-Item -Path "$ContextMenu_Basic_MSI\command" -Value $Command_For_MSI -force | out-null
																		Write_Log -Message_Type "INFO" -Message "Context menu for MSI has been added"																		
																	}

																write-progress -activity $Progress_Activity  -percentcomplete 55;


																If($Add_ZIP -eq $True)
																	{
																		# RUN ON ZIP
																		$ZIP_Shell_Registry_Key = "HKCR_SD:\CompressedFolder\Shell"
																		$ZIP_Basic_Run = "Extract ZIP in Sandbox"						
																		$ContextMenu_Basic_ZIP = "$ZIP_Shell_Registry_Key\$ZIP_Basic_Run"				

																		New-Item -Path $ZIP_Shell_Registry_Key -Name $ZiP_Basic_Run -force | out-null
																		New-Item -Path $ContextMenu_Basic_ZIP -Name "Command" -force | out-null
																		# Add Sandbox Icons
																		New-ItemProperty -Path $ContextMenu_Basic_ZIP -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null
																		$Command_For_ZIP = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type ZIP -LiteralPath "%V" -ScriptPath "%V"' 
																		# Set the command path
																		Set-Item -Path "$ContextMenu_Basic_ZIP\command" -Value $Command_For_ZIP -force | out-null
																		Write_Log -Message_Type "INFO" -Message "Context menu for ZIP has been added"																		
																	
																		write-progress -activity $Progress_Activity  -percentcomplete 65;

																		# RUN ON ZIP if WinRAR is installed
																		If(test-path "HKCR_SD:\WinRAR.ZIP\Shell")
																			{
																				$ZIP_WinRAR_Shell_Registry_Key = "HKCR_SD:\WinRAR.ZIP\Shell"
																				# $ZIP_WinRAR_Basic_Run = $Get_Language_File_Content.ZIP				
																				# $ContextMenu_Basic_ZIP_RAR = "$ZIP_WinRAR_Shell_Registry_Key\$ZIP_WinRAR_Basic_Run"
																				$ZIP_WinRAR_Basic_Run = "Extract ZIP in Sandbox"												
																				$ContextMenu_Basic_ZIP_RAR = "$ZIP_WinRAR_Shell_Registry_Key\$ZIP_WinRAR_Basic_Run"						
																				
																				New-Item -Path $ZIP_WinRAR_Shell_Registry_Key -Name $ZIP_WinRAR_Basic_Run -force | out-null
																				New-Item -Path $ContextMenu_Basic_ZIP_RAR -Name "Command" -force | out-null
																				# Add Sandbox Icons
																				New-ItemProperty -Path $ContextMenu_Basic_ZIP_RAR -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null
																				$Command_For_ZIP = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type ZIP -LiteralPath "%V" -ScriptPath "%V"' 
																				# Set the command path
																				Set-Item -Path "$ContextMenu_Basic_ZIP_RAR\command" -Value $Command_For_ZIP -force | out-null
																			}

																		$7z_Key_Label = "Extract 7z file in Sandbox"
																		# RUN ON 7z
																		$7z_Shell_Registry_Key = "HKCR_SD:\.7z"
																		If(test-path $7z_Shell_Registry_Key)
																			{
																				$Get_Default_Value = (Get-ItemProperty "HKCR_SD:\.7z")."(default)"
																				$Default_ZIP_Shell_Registry_Key = "HKCR_SD:\$Get_Default_Value\Shell"
																				If(test-path $Default_ZIP_Shell_Registry_Key)
																					{
																						$Default_ZIP_Key_Label_Path = "$Default_ZIP_Shell_Registry_Key\$7z_Key_Label"
																						$Default_ZIP_Command_Path = "$Default_ZIP_Key_Label_Path\Command"
																						$Command_for_Default_ZIP = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -windowstyle hidden -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type 7Z -LiteralPath "%V" -ScriptPath "%V"'
																						If(test-path $Default_ZIP_Shell_Registry_Key)
																						{
																							new-item $Default_ZIP_Key_Label_Path | out-null
																							new-item $Default_ZIP_Command_Path | out-null	
																							# Set the command path
																							Set-Item -Path $Default_ZIP_Command_Path -Value $Command_for_Default_ZIP -force | out-null	
																							# Add Sandbox Icons
																							New-ItemProperty -Path $Default_ZIP_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null			
																							Write_Log -Message_Type "INFO" -Message "Context menu for 7Z has been added"																									
																						}																						
																					}
																			}
																	}

																write-progress -activity $Progress_Activity  -percentcomplete 75;	
																		
																# RUN ON MSIX		
																If($Add_MSIX -eq $True)	
																	{
																		$MSIX_Key_Label = "Run MSIX file in Sandbox"	
																		$Command_for_MSIX = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -windowstyle hidden -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type MSIX -LiteralPath "%V" -ScriptPath "%V"'
																		
																		$MSIX_Shell_Registry_Key = "HKCR_SD:\.msix\OpenWithProgids"
																		If(test-path $MSIX_Shell_Registry_Key)
																			{
																				$Get_Default_Value = (Get-Item $MSIX_Shell_Registry_Key).Property
																				$MSIX_Shell_Registry = "HKCR_SD:\$Get_Default_Value\Shell"
																				If(test-path $MSIX_Shell_Registry)
																					{
																						$MSIX_Key_Label_Path = "$MSIX_Shell_Registry\$MSIX_Key_Label"
																						If(!(test-path $MSIX_Key_Label_Path))
																							{
																								$MSIX_Command_Path = "$MSIX_Key_Label_Path\Command"
																								new-item $MSIX_Key_Label_Path | out-null
																								new-item $MSIX_Command_Path | out-null	
																								# Set the command path
																								Set-Item -Path $MSIX_Command_Path -Value $Command_for_MSIX -force | out-null	
																								# Add Sandbox Icons
																								New-ItemProperty -Path $MSIX_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null			
																								Write_Log -Message_Type "INFO" -Message "Context menu for MSIX has been added"																									
																							}																							
																					}
																			}

																			# Modify value from HKCU
																			$Current_User_SID = (Get-ChildItem Registry::\HKEY_USERS | Where-Object { Test-Path "$($_.pspath)\Volatile Environment" } | ForEach-Object { (Get-ItemProperty "$($_.pspath)\Volatile Environment")}).PSParentPath.split("\")[-1]																			# RUN ON ISO
																			$HKCU_Classes = "Registry::HKEY_USERS\$Current_User_SID" + "_Classes"
																			If(test-path $HKCU_Classes)
																			{
																				$Default_MSIX_HKCU = "$HKCU_Classes\.msix"
																				$Get_Default_Value = (Get-Item "$Default_MSIX_HKCU\OpenWithProgids").Property
																				$Default_HKCU_MSIX_Shell_Registry_Key = "$HKCU_Classes\$Get_Default_Value\Shell"
																				If(test-path $Default_HKCU_MSIX_Shell_Registry_Key)
																					{
																						$MSIX_HKCU_Key_Label_Path = "$Default_HKCU_MSIX_Shell_Registry_Key\$MSIX_Key_Label"																					
																						$HKCU_MSIX_Command_Path = "$MSIX_HKCU_Key_Label_Path\Command"																					
																						new-item $MSIX_HKCU_Key_Label_Path | out-null
																						new-item $HKCU_MSIX_Command_Path | out-null	
																						# Set the command path
																						Set-Item -Path $HKCU_MSIX_Command_Path -Value $Command_for_MSIX -force | out-null	
																						# Add Sandbox Icons
																						New-ItemProperty -Path $MSIX_HKCU_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null			
																						Write_Log -Message_Type "INFO" -Message "Context menu for ISO has been added"																																													
																																						
																					}																																												
																			}																			
																	}
																	
																write-progress -activity $Progress_Activity  -percentcomplete 75;	
																	

																If($Add_Folder -eq $True)
																	{
																		# Share this folder - Inside the folder
																		$Folder_Inside_Shell_Registry_Key = "HKCR_SD:\Directory\Background\shell"
																		$Folder_Inside_Basic_Run = "Share this folder in a Sandbox"	
																		# $Folder_Inside_Basic_Run = $Get_Language_File_Content.Folder										
																		# $ContextMenu_Folder_Inside = "$Folder_Inside_Shell_Registry_Key\$Folder_Inside_Basic_Run"
																		$ContextMenu_Folder_Inside = "$Folder_Inside_Shell_Registry_Key\Share this folder in a Sandbox"

																		New-Item -Path $Folder_Inside_Shell_Registry_Key -Name $Folder_Inside_Basic_Run -force | out-null
																		New-Item -Path $ContextMenu_Folder_Inside -Name "Command" -force | out-null
																		# Add Sandbox Icons
																		New-ItemProperty -Path $ContextMenu_Folder_Inside -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null
																		$Command_For_Folder_Inside = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type Folder_Inside -LiteralPath "%V" -ScriptPath "%V"' 
																		# Set the command path
																		Set-Item -Path "$ContextMenu_Folder_Inside\command" -Value $Command_For_Folder_Inside -force | out-null
																		Write_Log -Message_Type "INFO" -Message "Context menus for folder have been added"									


																		write-progress -activity $Progress_Activity  -percentcomplete 85;	


																		# Share this folder - Right-click on the folder
																		$Folder_On_Shell_Registry_Key = "HKCR_SD:\Directory\shell"
																		# $Folder_On_Run = $Get_Language_File_Content.Folder		
																		$Folder_On_Run = "Share this folder in a Sandbox"										
																		# $ContextMenu_Folder_On = "$Folder_On_Shell_Registry_Key\$Folder_On_Run"
																		$ContextMenu_Folder_On = "$Folder_On_Shell_Registry_Key\Share this folder in a Sandbox"

																		New-Item -Path $Folder_On_Shell_Registry_Key -Name $Folder_On_Run -force | out-null
																		New-Item -Path $ContextMenu_Folder_On -Name "Command" -force | out-null
																		# Add Sandbox Icons
																		New-ItemProperty -Path $ContextMenu_Folder_On -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null
																		$Command_For_Folder_On = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type Folder_On -LiteralPath "%V" -ScriptPath "%V"' 
																		# Set the command path
																		Set-Item -Path "$ContextMenu_Folder_On\command" -Value $Command_For_Folder_On -force | out-null
																	}

																If($List_Drive -ne $null){Remove-PSDrive $List_Drive}

																write-progress -activity $Progress_Activity  -percentcomplete 100;	
																copy-item $Log_File $Destination_folder -Force																
															}													
													}
												Else
													{
														Write_Log -Message_Type "ERROR" -Message "File RunInSandbox.ps1 is missing"														
														[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
														[System.Windows.Forms.MessageBox]::Show("File RunInSandbox.ps1 is missing !!!")	
													}													
											}
								
									}											
							}
						Else
							{
								Write_Log -Message_Type "ERROR" -Message "Some contents are missing"														
								[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
								[System.Windows.Forms.MessageBox]::Show("It seems you don't have dowloaded all the folder structure !!!")	
							}
					}
				Else
					{
						Write_Log -Message_Type "INFO" -Message "Sources folder is missing"																	
						[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
						[System.Windows.Forms.MessageBox]::Show("It seems you don't have dowloaded all the folder structure.`nThe folder Sources is missing !!!")	
					}
			}
	}