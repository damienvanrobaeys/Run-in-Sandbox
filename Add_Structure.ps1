#***************************************************************************************************************
# Author: Damien VAN ROBAEYS
# Website: http://www.systanddeploy.com
# Twitter: https://twitter.com/syst_and_deploy
#***************************************************************************************************************
$TEMP_Folder = $env:temp
$Log_File = "$TEMP_Folder\Enable_RunInSandbox.log"
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
		Add-Content $Log_File  "$MyDate - $Message_Type : $Message"			
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
						reg export "HKEY_CLASSES_ROOT\$Reg_Path" $Backup_Path | out-null
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
								
						$Check_Sources_Files_Count = (get-childitem "$Current_Folder\Sources" -recurse).count
						If($Check_Sources_Files_Count -eq 19)	
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
													
												If(test-path "$ProgData\Run_in_Sandbox\RunInSandbox.ps1")
													{													
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

																Export_Reg_Config -Reg_Path "exefile" -Backup_Path "$Destination_folder\Backup_HKRoot_EXEFile.reg"
																Export_Reg_Config -Reg_Path "Microsoft.PowerShellScript.1" -Backup_Path "$Destination_folder\Backup_HKRoot_PowerShellScript.reg"
																Export_Reg_Config -Reg_Path "VBSFile" -Backup_Path "$Destination_folder\Backup_HKRoot_VBSFile.reg"
																Export_Reg_Config -Reg_Path "Msi.Package" -Backup_Path "$Destination_folder\Backup_HKRoot_Msi.reg"
																Export_Reg_Config -Reg_Path "CompressedFolder" -Backup_Path "$Destination_folder\Backup_HKRoot_CompressedFolder.reg"
																Export_Reg_Config -Reg_Path "WinRAR.ZIP" -Backup_Path "$Destination_folder\Backup_HKRoot_WinRAR.reg"
																Export_Reg_Config -Reg_Path "Directory" -Backup_Path "$Destination_folder\Backup_HKRoot_Directory.reg"						

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
																		# RUN ON PS1
																		$PS1_Shell_Registry_Key = "HKCR_SD:\Microsoft.PowerShellScript.1\Shell"
																		# $PS1_Basic_Run = $Get_Language_File_Content.PowerShell.Basic
																		# $PS1_Parameter_Run = $Get_Language_File_Content.PowerShell.Parameters
																		# $ContextMenu_Basic_PS1 = "$PS1_Shell_Registry_Key\$PS1_Basic_Run"
																		# $ContextMenu_Parameters_PS1 = "$PS1_Shell_Registry_Key\$PS1_Parameter_Run"
																		$PS1_Basic_Run = "Run the PS1 in Sandbox"
																		$PS1_Parameter_Run = "Run the PS1 in Sandbox with parameters"						
																		$ContextMenu_Basic_PS1 = "$PS1_Shell_Registry_Key\Run the PS1 in Sandbox"
																		$ContextMenu_Parameters_PS1 = "$PS1_Shell_Registry_Key\Run the PS1 in Sandbox with parameters"				
																		
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
																		Write_Log -Message_Type "INFO" -Message "Context menus for PS1 have been added"																		
																	}
																	
																write-progress -activity $Progress_Activity  -percentcomplete 25;


																If($Add_Intunewin -eq $True)
																	{
																		# RUN ON INTUNEWIN
																		$Intunewin_Shell_Registry_Key = "HKCR_SD:\.intunewin"
																		$Intunewin_Key_Label = "Test the intunewin in Sandbox"
																		$Intunewin_Key_Label_Path = "$Intunewin_Shell_Registry_Key\Shell\$Intunewin_Key_Label"
																		$Intunewin_Command_Path = "$Intunewin_Key_Label_Path\Command"
																		$Command_for_Intunewin = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -windowstyle hidden -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type Intunewin -LiteralPath "%V" -ScriptPath "%V"'
																		If(!(test-path $Intunewin_Shell_Registry_Key))
																		{
																			new-item $Intunewin_Shell_Registry_Key | out-null
																			new-item "$Intunewin_Shell_Registry_Key\Shell" | out-null
																			new-item $Intunewin_Key_Label_Path | out-null
																			new-item $Intunewin_Command_Path | out-null	
																			# Set the command path
																			Set-Item -Path $Intunewin_Command_Path -Value $Command_for_Intunewin -force | out-null	
																			# Add Sandbox Icons
																			New-ItemProperty -Path $Intunewin_Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null			
																			Write_Log -Message_Type "INFO" -Message "Context menus for PS1 have been added"																									
																		}																	
																	}
																
																write-progress -activity $Progress_Activity  -percentcomplete 30;

																If($Add_VBS -eq $True)
																	{
																		# RUN ON VBS
																		$VBS_Shell_Registry_Key = "HKCR_SD:\VBSFile\Shell"
																		# $VBS_Basic_Run = $Get_Language_File_Content.VBS.Basic
																		# $VBS_Parameter_Run = $Get_Language_File_Content.VBS.Parameters
																		# $ContextMenu_Basic_VBS = "$VBS_Shell_Registry_Key\$VBS_Basic_Run"
																		# $ContextMenu_Parameters_VBS = "$VBS_Shell_Registry_Key\$VBS_Parameter_Run"
																		
																		$VBS_Basic_Run = "Run the VBS in Sandbox"
																		$VBS_Parameter_Run = "Run the VBS in Sandbox with parameters"					
																		
																		$ContextMenu_Basic_VBS = "$VBS_Shell_Registry_Key\Run the VBS in Sandbox"
																		$ContextMenu_Parameters_VBS = "$VBS_Shell_Registry_Key\Run the VBS in Sandbox with parameters"				

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

																write-progress -activity $Progress_Activity  -percentcomplete 40;

																If($Add_EXE -eq $True)
																	{
																		# RUN ON EXE
																		$EXE_Shell_Registry_Key = "HKCR_SD:\exefile\Shell"
																		# $EXE_Basic_Run = $Get_Language_File_Content.EXE
																		# $ContextMenu_Basic_EXE = "$EXE_Shell_Registry_Key\$EXE_Basic_Run"
																		$EXE_Basic_Run = "Run the EXE in Sandbox"					
																		$ContextMenu_Basic_EXE = "$EXE_Shell_Registry_Key\Run the EXE in Sandbox"				

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
																		# $MSI_Basic_Run = $Get_Language_File_Content.MSI
																		# $ContextMenu_Basic_MSI = "$MSI_Shell_Registry_Key\$MSI_Basic_Run"
																		$MSI_Basic_Run = "Run the MSI in Sandbox"					
																		$ContextMenu_Basic_MSI = "$MSI_Shell_Registry_Key\Run the MSI in Sandbox"				

																		New-Item -Path $MSI_Shell_Registry_Key -Name $MSI_Basic_Run -force | out-null
																		New-Item -Path $ContextMenu_Basic_MSI -Name "Command" -force | out-null
																		# Add Sandbox Icons
																		New-ItemProperty -Path $ContextMenu_Basic_MSI -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null
																		$Command_For_MSI = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type MSI -LiteralPath "%V" -ScriptPath "%V"' 
																		# Set the command path
																		Set-Item -Path "$ContextMenu_Basic_MSI\command" -Value $Command_For_MSI -force | out-null
																		Write_Log -Message_Type "INFO" -Message "Context menus for MSI have been added"																		
																	}

																write-progress -activity $Progress_Activity  -percentcomplete 60;


																If($Add_ZIP -eq $True)
																	{
																		# RUN ON ZIP
																		$ZIP_Shell_Registry_Key = "HKCR_SD:\CompressedFolder\Shell"
																		# $ZIP_Basic_Run = $Get_Language_File_Content.ZIP
																		# $ContextMenu_Basic_ZIP = "$ZIP_Shell_Registry_Key\$ZiP_Basic_Run"
																		$ZIP_Basic_Run = "Extract the ZIP in Sandbox"						
																		$ContextMenu_Basic_ZIP = "$ZIP_Shell_Registry_Key\Extract the ZIP in Sandbox"				

																		New-Item -Path $ZIP_Shell_Registry_Key -Name $ZiP_Basic_Run -force | out-null
																		New-Item -Path $ContextMenu_Basic_ZIP -Name "Command" -force | out-null
																		# Add Sandbox Icons
																		New-ItemProperty -Path $ContextMenu_Basic_ZIP -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null
																		$Command_For_ZIP = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type ZIP -LiteralPath "%V" -ScriptPath "%V"' 
																		# Set the command path
																		Set-Item -Path "$ContextMenu_Basic_ZIP\command" -Value $Command_For_ZIP -force | out-null
																		Write_Log -Message_Type "INFO" -Message "Context menus for ZIP have been added"																		
																	
																		write-progress -activity $Progress_Activity  -percentcomplete 70;

																		# RUN ON ZIP if WinRAR is installed
																		If(test-path "HKCR_SD:\WinRAR.ZIP\Shell")
																			{
																				$ZIP_WinRAR_Shell_Registry_Key = "HKCR_SD:\WinRAR.ZIP\Shell"
																				# $ZIP_WinRAR_Basic_Run = $Get_Language_File_Content.ZIP				
																				# $ContextMenu_Basic_ZIP_RAR = "$ZIP_WinRAR_Shell_Registry_Key\$ZIP_WinRAR_Basic_Run"
																				$ZIP_WinRAR_Basic_Run = "Extract the ZIP in Sandbox"												
																				$ContextMenu_Basic_ZIP_RAR = "$ZIP_WinRAR_Shell_Registry_Key\Extract the ZIP in Sandbox"						
																				
																				New-Item -Path $ZIP_WinRAR_Shell_Registry_Key -Name $ZIP_WinRAR_Basic_Run -force | out-null
																				New-Item -Path $ContextMenu_Basic_ZIP_RAR -Name "Command" -force | out-null
																				# Add Sandbox Icons
																				New-ItemProperty -Path $ContextMenu_Basic_ZIP_RAR -Name "icon" -PropertyType String -Value $Sandbox_Icon | out-null
																				$Command_For_ZIP = 'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -executionpolicy bypass -sta -file C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -NoExit -Command Set-Location -Type ZIP -LiteralPath "%V" -ScriptPath "%V"' 
																				# Set the command path
																				Set-Item -Path "$ContextMenu_Basic_ZIP_RAR\command" -Value $Command_For_ZIP -force | out-null
																			}																	
																	}

																write-progress -activity $Progress_Activity  -percentcomplete 80;	
																
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


																		write-progress -activity $Progress_Activity  -percentcomplete 90;	


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
