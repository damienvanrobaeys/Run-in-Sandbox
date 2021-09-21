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
		
		$List_Drive = get-psdrive | where {$_.Name -eq "HKCR_SD"}
		If($List_Drive -ne $null){Remove-PSDrive $List_Drive}
		New-PSDrive -PSProvider registry -Root HKEY_CLASSES_ROOT -Name HKCR_SD | out-null

		If($Add_PS1 -eq $True)
			{
				# REMOVE RUN ON PS1
				write-host "Removing context menu for PS1"				
				$PS1_Shell_Registry_Key = "HKCR_SD:\Microsoft.PowerShellScript.1\Shell"
				$PS1_Basic_Run = "Run the PS1 in Sandbox"
				$PS1_Parameter_Run = "Run the PS1 in Sandbox with parameters"					
				Remove_Reg_Item -Reg_Path "$PS1_Shell_Registry_Key\$PS1_Basic_Run"
				Remove_Reg_Item -Reg_Path "$PS1_Shell_Registry_Key\$PS1_Parameter_Run"			
			}
			
		If($Add_Reg -eq $True)
			{
				# REMOVE RUN ON REG
				write-host "Removing context menu for REG"				
				$Reg_Shell_Registry_Key = "HKCR_SD:\regfile\Shell"
				$Reg_Key_Label = "Test the reg file in Sandbox"
				Remove_Reg_Item -Reg_Path "$REG_Shell_Registry_Key\$Reg_Key_Label"
			}					
			
		If($Add_EXE -eq $True)
			{
				# REMOVE RUN ON EXE
				write-host "Removing context menu for PS1"							
				$EXE_Shell_Registry_Key = "HKCR_SD:\exefile\Shell"
				$EXE_Basic_Run = "Run the EXE in Sandbox"
				Remove_Reg_Item -Reg_Path "$EXE_Shell_Registry_Key\$EXE_Basic_Run"			
			}

		If($Add_MSI -eq $True)
			{
				# RUN ON MSI
				write-host "Removing context menu for MSI"				
				$MSI_Shell_Registry_Key = "HKCR_SD:\Msi.Package\Shell"
				$MSI_Basic_Run = "Run the MSI in Sandbox"	
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
				$VBS_Basic_Run = "Run the VBS in Sandbox"
				$VBS_Parameter_Run = "Run the VBS in Sandbox with parameters"			
				Remove_Reg_Item -Reg_Path "$VBS_Shell_Registry_Key\$VBS_Basic_Run"
				Remove_Reg_Item -Reg_Path "$VBS_Shell_Registry_Key\$VBS_Parameter_Run"				
			}

		If($Add_ZIP -eq $True)
			{
				write-host "Removing context menu for ZIP"			
				# RUN ON ZIP				
				$ZIP_Shell_Registry_Key = "HKCR_SD:\CompressedFolder\Shell"
				$ZIP_Basic_Run = "Extract the ZIP in Sandbox"	
				Remove_Reg_Item -Reg_Path "$ZIP_Shell_Registry_Key\$ZIP_Basic_Run"		

				# RUN ON ZIP if WinRAR is installed
				If(test-path "HKCR_SD:\WinRAR.ZIP\Shell")
					{
						$ZIP_WinRAR_Shell_Registry_Key = "HKCR_SD:\WinRAR.ZIP\Shell"
						$ZIP_WinRAR_Basic_Run = "Extract the ZIP in Sandbox"					
						Remove_Reg_Item -Reg_Path "$ZIP_WinRAR_Shell_Registry_Key\$ZIP_WinRAR_Basic_Run"								
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