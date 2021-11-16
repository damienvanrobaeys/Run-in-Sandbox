#***************************************************************************************************************
# Author: Damien VAN ROBAEYS
# Website: http://www.systanddeploy.com
# Twitter: https://twitter.com/syst_and_deploy
# Purpose: This script will remove context menus added to run quickly files in Windows Sandbox
#***************************************************************************************************************

Function Write_Log
	{
		param(
		$Message_Type,	
		$Message
		)
		
		$MyDate = "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)		
		write-host "$MyDate - $Message_Type : $Message"			
	}

Function Remove_Reg_Item
	{
		param(
		$Reg_Path		
		)
		
		Write_Log -Message_Type "INFO" -Message "Removing registry path: $Reg_Path"		
				
		If(test-path $Reg_Path)
			{
				Write_Log -Message_Type "SUCCESS" -Message "Following registry path exists: $Reg_Path"												
				Try
					{
						Remove-Item -Path $Reg_Path -Recurse						
						Write_Log -Message_Type "SUCCESS" -Message "$Reg_Path has been removed"									
					}
				Catch
					{
						Write_Log -Message_Type "ERROR" -Message "$Reg_Path has not been removed"									
					}			
			}
		Else
			{
				Write_Log -Message_Type "INFO" -Message "Can not find registry path: $Reg_Path"															
			}				
	}	


$Current_Folder = split-path $MyInvocation.MyCommand.Path
$ProgData = $env:ProgramData
$Sandbox_Folder = "$ProgData\Run_in_Sandbox"
If(test-path $Sandbox_Folder)
	{
		# $XML_Config = "$Sandbox_Folder\Sandbox_Config.xml"
		# $Get_XML_Content = [xml] (Get-Content $XML_Config)
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

		$XML_Config = "$Sandbox_Folder\Sandbox_Config.xml"																							
		$Get_XML_Content = [xml] (Get-Content $XML_Config)
		
		# Check which context menu should be remvoved
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
		
		$List_Drive = get-psdrive | where {$_.Name -eq "HKCR_SD"}
		If($List_Drive -ne $null){Remove-PSDrive $List_Drive}
		New-PSDrive -PSProvider registry -Root HKEY_CLASSES_ROOT -Name HKCR_SD | out-null

		If($Add_PS1 -eq $True)
			{
				# REMOVE RUN ON PS1
				write-host "Removing context menu for PS1"				

				$PS1_Main_Menu = "Run PS1 in Sandbox"																		
				$Windows_Version = (Get-WmiObject -class Win32_OperatingSystem).Caption
				If($Windows_Version -like "*Windows 10*")
					{
						$PS1_Shell_Registry_Key = "HKCR_SD:\Microsoft.PowerShellScript.1\Shell"
						$PS1_Basic_Run = "Run PS1 in Sandbox"
						Remove_Reg_Item -Reg_Path "$PS1_Shell_Registry_Key\$PS1_Basic_Run"
					}
				ElseIf($Windows_Version -like "*Windows 11*")
					{
						$Current_User_SID = (Get-ChildItem Registry::\HKEY_USERS | Where-Object { Test-Path "$($_.pspath)\Volatile Environment" } | ForEach-Object { (Get-ItemProperty "$($_.pspath)\Volatile Environment")}).PSParentPath.split("\")[-1]																			# RUN ON ISO
						$HKCU_Classes = "Registry::HKEY_USERS\$Current_User_SID" + "_Classes"
						If(test-path $HKCU_Classes)
						{
							$Default_PS1_HKCU = "$HKCU_Classes\.ps1"
							$Get_Default_Value = (Get-Item "$Default_PS1_HKCU\rOpenWithProgids").Property

							$Default_HKCU_PS1_Shell_Registry_Key = "$HKCU_Classes\$Get_Default_Value\Shell"
							If(test-path $Default_HKCU_PS1_Shell_Registry_Key)
								{
									$Main_Menu_Path = "$Default_HKCU_PS1_Shell_Registry_Key\$PS1_Main_Menu"
									Remove_Reg_Item -Reg_Path "$Main_Menu_Path"									
								}																																												
						}
					}				
			}
			
		If($Add_Reg -eq $True)
			{
				# REMOVE RUN ON REG
				write-host "Removing context menu for REG"				
				$Reg_Shell_Registry_Key = "HKCR_SD:\regfile\Shell"
				$Reg_Key_Label = "Test reg file in Sandbox"
				Remove_Reg_Item -Reg_Path "$REG_Shell_Registry_Key\$Reg_Key_Label"
			}	

		If($Add_ISO -eq $True)
			{
				$ISO_Key_Label = "Extract ISO file in Sandbox"				
			
				# REMOVE RUN ON REG from HKCR
				write-host "Removing context menu for ISO"		
				$ISO_Shell_Registry_Key = "HKCR_SD:\Windows.IsoFile\Shell"
				Remove_Reg_Item -Reg_Path "$ISO_Shell_Registry_Key\$ISO_Key_Label"
				
				# REMOVE RUN ON REG from HKCU if 7zip exists
				$ISO_Shell_HKCU_Registry_Key = "Registry::HKEY_USERS\$Current_User_SID"				
				$HKCU_Classes = "$ISO_Shell_HKCU_Registry_Key\SOFTWARE\Classes"
				$Default_ISO_HKCU = "$HKCU_Classes\.iso"
				If(test-path $Default_ISO_HKCU)
					{
						$Get_Default_Value = (Get-ItemProperty $Default_ISO_HKCU)."(default)"
						If($Get_Default_Value -eq "7-Zip.iso")
							{
								$Default_HKCU_ISO_Shell_Registry_Key = "$HKCU_Classes\$Get_Default_Value\Shell"
								$ISO_HKCU_Key_Label_Path = "$Default_HKCU_ISO_Shell_Registry_Key\$ISO_Key_Label"
								If(test-path $ISO_HKCU_Key_Label_Path)
									{
										Remove_Reg_Item -Reg_Path $ISO_HKCU_Key_Label_Path
									}
							}					
					}
			}	
			
			
		If($Add_MSIX -eq $True)	
			{
				$MSIX_Key_Label = "Run MSIX file in Sandbox"																					
				# REMOVE RUN ON REG from HKCR
				$MSIX_Shell_Registry_Key = "HKCR_SD:\.msix\OpenWithProgids"
				If(test-path $MSIX_Shell_Registry_Key)
					{
						$Get_Default_Value = (Get-Item $MSIX_Shell_Registry_Key).Property
						$MSIX_Shell_Registry = "HKCR_SD:\$Get_Default_Value\Shell"
						If(test-path $MSIX_Shell_Registry)
							{
								$MSIX_Key_Label_Path = "$MSIX_Shell_Registry\$MSIX_Key_Label"
								If(test-path $MSIX_Key_Label_Path)
									{
										Remove_Reg_Item -Reg_Path "$MSIX_Key_Label_Path"														
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
						$MSIX_HKCU_Key_Label_Path = "$Default_HKCU_MSIX_Shell_Registry_Key\$MSIX_Key_Label"						
						If(test-path $MSIX_HKCU_Key_Label_Path)
							{
								Remove_Reg_Item -Reg_Path "$MSIX_HKCU_Key_Label_Path"
							}																																												
					}	
			}			

		If($Add_PPKG -eq $True)
			{
				# REMOVE RUN ON PPKG
				write-host "Removing context menu for PPKG"		
				$PPKG_Shell_Registry_Key = "HKCR_SD:\Microsoft.ProvTool.Provisioning.1\Shell"
				$PPKG_Key_Label = "Run PPKG file in Sandbox"				
				Remove_Reg_Item -Reg_Path "$PPKG_Shell_Registry_Key\$PPKG_Key_Label"
			}	

		If($Add_HTML -eq $True)
			{
				$HTML_Key_Label = "Run this web link in Sandbox"
				
				# RUN ON HTML for Edge
				$HTML_Edge_Shell_Registry_Key = "HKCR_SD:\MSEdgeHTM\Shell"
				Remove_Reg_Item -Reg_Path "$HTML_Edge_Shell_Registry_Key\$HTML_Key_Label"																								
					
				# RUN ON HTML for Chrome
				$HTML_Chrome_Shell_Registry_Key = "HKCR_SD:\ChromeHTML\Shell"
				Remove_Reg_Item -Reg_Path "$HTML_Chrome_Shell_Registry_Key\$HTML_Key_Label"																																												

				# RUN ON HTML for IE
				$HTML_IE_Shell_Registry_Key = "HKCR_SD:\IE.AssocFile.HTM\Shell"
				Remove_Reg_Item -Reg_Path "$HTML_IE_Shell_Registry_Key\$HTML_Key_Label"		

				# RUN ON URL
				$URL_Shell_Registry_Key = "HKCR_SD:\IE.AssocFile.URL\Shell"
				$URL_Key_Label_Path = "Run this URL in Sandbox"
				# $URL_Key_Label_Path = "Run this web link in Sandbox"				
				Remove_Reg_Item -Reg_Path "$URL_Shell_Registry_Key\$URL_Key_Label_Path"						
			}				
			
		If($Add_EXE -eq $True)
			{
				# REMOVE RUN ON EXE
				write-host "Removing context menu for PS1"							
				$EXE_Shell_Registry_Key = "HKCR_SD:\exefile\Shell"
				$EXE_Basic_Run = "Run EXE in Sandbox"
				Remove_Reg_Item -Reg_Path "$EXE_Shell_Registry_Key\$EXE_Basic_Run"			
			}

		If($Add_MSI -eq $True)
			{
				# RUN ON MSI
				write-host "Removing context menu for MSI"				
				$MSI_Shell_Registry_Key = "HKCR_SD:\Msi.Package\Shell"
				$MSI_Basic_Run = "Run MSI in Sandbox"	
				Remove_Reg_Item -Reg_Path "$MSI_Shell_Registry_Key\$MSI_Basic_Run"				
			}
			
		If($Add_Folder -eq $True)
			{
				write-host "Removing context menu for folder"			
				# Share this folder - Inside the folder
				$Folder_Inside_Shell_Registry_Key = "HKCR_SD:\Directory\Background\shell"
				$Folder_Inside_Basic_Run = "Share this folder in a Sandbox"						
				Remove_Reg_Item -Reg_Path "$Folder_Inside_Shell_Registry_Key\$Folder_Inside_Basic_Run"										

				# Share this folder - Right-click on the folder
				$Folder_On_Shell_Registry_Key = "HKCR_SD:\Directory\shell"
				$Folder_On_Run = "Share this folder in a Sandbox"						
				Remove_Reg_Item -Reg_Path "$Folder_On_Shell_Registry_Key\$Folder_On_Run"				
			}

		If($Add_Intunewin -eq $True)
			{
				# RUN ON Intunewin
				write-host "Removing context menu for intunewin"				
				Remove_Reg_Item -Reg_Path "HKCR_SD:\.intunewin"								
			}
			
		If($Add_MultipleApp -eq $True)
			{
				# RUN ON multiple app context menu
				write-host "Removing context menu for multiple app"				
				Remove_Reg_Item -Reg_Path "HKCR_SD:\.sdbapp"								
			}			
			
		If($Add_VBS -eq $True)
			{
				# REMOVE RUN ON VBS
				write-host "Removing context menu for VBS"				
				$VBS_Shell_Registry_Key = "HKCR_SD:\VBSFile\Shell"
				$VBS_Basic_Run = "Run VBS in Sandbox"
				$VBS_Parameter_Run = "Run VBS in Sandbox with parameters"			
				Remove_Reg_Item -Reg_Path "$VBS_Shell_Registry_Key\$VBS_Basic_Run"
				Remove_Reg_Item -Reg_Path "$VBS_Shell_Registry_Key\$VBS_Parameter_Run"				
			}

		If($Add_ZIP -eq $True)
			{
				write-host "Removing context menu for ZIP"			
				# RUN ON ZIP				
				$ZIP_Shell_Registry_Key = "HKCR_SD:\CompressedFolder\Shell"
				$ZIP_Basic_Run = "Extract ZIP in Sandbox"	
				Remove_Reg_Item -Reg_Path "$ZIP_Shell_Registry_Key\$ZIP_Basic_Run"		

				# RUN ON ZIP if WinRAR is installed
				If(test-path "HKCR_SD:\WinRAR.ZIP\Shell")
					{
						$ZIP_WinRAR_Shell_Registry_Key = "HKCR_SD:\WinRAR.ZIP\Shell"
						# $ZIP_WinRAR_Basic_Run = "Extract ZIP in Sandbox"					
						Remove_Reg_Item -Reg_Path "$ZIP_WinRAR_Shell_Registry_Key\$ZIP_Basic_Run"								
					}

				# REMOVE RUN ON 7Z
				$7z_Key_Label = "Extract 7z file in Sandbox"
				$7z_Shell_Registry_Key = "HKCR_SD:\.7z"
				If(test-path $7z_Shell_Registry_Key)
					{
						$Get_Default_Value = (Get-ItemProperty "HKCR_SD:\.7z")."(default)"
						$Default_ZIP_Shell_Registry_Key = "HKCR_SD:\$Get_Default_Value\Shell"
						If(test-path $Default_ZIP_Shell_Registry_Key)
							{
								If(test-path $Default_ZIP_Shell_Registry_Key)
								{
									Remove_Reg_Item -Reg_Path "$Default_ZIP_Shell_Registry_Key\$7z_Key_Label"
								}																						
							}
					}					
			}

		If($List_Drive -ne $null){Remove-PSDrive $List_Drive}

		Remove-item $Sandbox_Folder -recurse -force	
	}
Else
	{
		[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
		[System.Windows.Forms.MessageBox]::Show("Can not find the folder $Sandbox_Folder")	
	}