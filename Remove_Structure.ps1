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


		$List_Drive = get-psdrive | where {$_.Name -eq "HKCR_SD"}
		If($List_Drive -ne $null){Remove-PSDrive $List_Drive}
		New-PSDrive -PSProvider registry -Root HKEY_CLASSES_ROOT -Name HKCR_SD | out-null


		# REMOVE RUN ON PS1
		$PS1_Shell_Registry_Key = "HKCR_SD:\Microsoft.PowerShellScript.1\Shell"
		# $PS1_Basic_Run = $Get_Language_File_Content.PowerShell.Basic
		# $PS1_Parameter_Run = $Get_Language_File_Content.PowerShell.Parameters		
		$PS1_Basic_Run = "Run the PS1 in Sandbox"
		$PS1_Parameter_Run = "Run the PS1 in Sandbox with parameters"			
		# Remove-Item -Path "$PS1_Shell_Registry_Key\$PS1_Basic_Run" -Recurse
		# Remove-Item -Path "$PS1_Shell_Registry_Key\$PS1_Parameter_Run" -Recurse
		
		Remove_Reg_Item -Reg_Path "$PS1_Shell_Registry_Key\$PS1_Basic_Run"
		Remove_Reg_Item -Reg_Path "$PS1_Shell_Registry_Key\$PS1_Parameter_Run"
		
		
		# REMOVE RUN ON VBS
		$VBS_Shell_Registry_Key = "HKCR_SD:\VBSFile\Shell"
		# $VBS_Basic_Run = $Get_Language_File_Content.VBS.Basic
		# $VBS_Parameter_Run = $Get_Language_File_Content.VBS.Parameters		
		$VBS_Basic_Run = "Run the VBS in Sandbox"
		$VBS_Parameter_Run = "Run the VBS in Sandbox with parameters"	
		# Remove-Item -Path "$VBS_Shell_Registry_Key\$VBS_Basic_Run" -Recurse
		# Remove-Item -Path "$VBS_Shell_Registry_Key\$VBS_Parameter_Run" -Recurse
		
		Remove_Reg_Item -Reg_Path "$VBS_Shell_Registry_Key\$VBS_Basic_Run"
		Remove_Reg_Item -Reg_Path "$VBS_Shell_Registry_Key\$VBS_Parameter_Run"		


		# REMOVE RUN ON EXE
		$EXE_Shell_Registry_Key = "HKCR_SD:\exefile\Shell"
		# $EXE_Basic_Run = $Get_Language_File_Content.EXE
		$EXE_Basic_Run = "Run the EXE in Sandbox"
		# Remove-Item -Path "$EXE_Shell_Registry_Key\$EXE_Basic_Run" -Recurse
		Remove_Reg_Item -Reg_Path "$EXE_Shell_Registry_Key\$EXE_Basic_Run"


		# RUN ON MSI
		$MSI_Shell_Registry_Key = "HKCR_SD:\Msi.Package\Shell"
		# $MSI_Basic_Run = $Get_Language_File_Content.MSI
		$MSI_Basic_Run = "Run the MSI in Sandbox"	
		# Remove-Item -Path "$MSI_Shell_Registry_Key\$MSI_Basic_Run" -Recurse
		Remove_Reg_Item -Reg_Path "$MSI_Shell_Registry_Key\$MSI_Basic_Run"		


		# RUN ON ZIP
		$ZIP_Shell_Registry_Key = "HKCR_SD:\CompressedFolder\Shell"
		# $ZIP_Basic_Run = $Get_Language_File_Content.ZIP
		$ZIP_Basic_Run = "Extract the ZIP in Sandbox"	
		# Remove-Item -Path "$ZIP_Shell_Registry_Key\$ZIP_Basic_Run" -Recurse
		Remove_Reg_Item -Reg_Path "$ZIP_Shell_Registry_Key\$ZIP_Basic_Run"				


		# RUN ON ZIP if WinRAR is installed
		If(test-path "HKCR_SD:\WinRAR.ZIP\Shell")
			{
				$ZIP_WinRAR_Shell_Registry_Key = "HKCR_SD:\WinRAR.ZIP\Shell"
				# $ZIP_WinRAR_Basic_Run = $Get_Language_File_Content.ZIP	
				$ZIP_WinRAR_Basic_Run = "Extract the ZIP in Sandbox"					
				# Remove-Item -Path "$ZIP_WinRAR_Shell_Registry_Key\$ZIP_WinRAR_Basic_Run" -Recurse
				Remove_Reg_Item -Reg_Path "$ZIP_WinRAR_Shell_Registry_Key\$ZIP_WinRAR_Basic_Run"								
			}
			
			
		# Share this folder - Inside the folder
		$Folder_Inside_Shell_Registry_Key = "HKCR_SD:\Directory\Background\shell"
		# $Folder_Inside_Basic_Run = $Get_Language_File_Content.Folder	
		$Folder_Inside_Basic_Run = "Share this folder in a Sandbox"						
		# Remove-Item -Path "$Folder_Inside_Shell_Registry_Key\$Folder_Inside_Basic_Run" -Recurse
		Remove_Reg_Item -Reg_Path "$Folder_Inside_Shell_Registry_Key\$Folder_Inside_Basic_Run"										


		# Share this folder - Right-click on the folder
		$Folder_On_Shell_Registry_Key = "HKCR_SD:\Directory\shell"
		# $Folder_On_Run = $Get_Language_File_Content.Folder	
		$Folder_On_Run = "Share this folder in a Sandbox"						
		# Remove-Item -Path "$Folder_On_Shell_Registry_Key\$Folder_On_Run" -Recurse	
		Remove_Reg_Item -Reg_Path "$Folder_On_Shell_Registry_Key\$Folder_On_Run"												

		If($List_Drive -ne $null){Remove-PSDrive $List_Drive}

		Remove-item $Sandbox_Folder -recurse -force	
	}
Else
	{
		[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
		[System.Windows.Forms.MessageBox]::Show("Can not find the folder $Sandbox_Folder")	
	}