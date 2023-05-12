$Desktop = "C:\Users\WDAGUtilityAccount\Desktop"
$Sandbox_Root_Path = "C:\Run_in_Sandbox"
$App_Bundle_File = "$Sandbox_Root_Path\App_Bundle.sdbapp"
$Get_Apps_to_install = [xml](Get-Content $App_Bundle_File)
$Apps_to_install = $Get_Apps_to_install.Applications.Application
ForEach ($App in $Apps_to_install) {
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

	$App_Path = $App.Path
	$App_File = $App.File
	$App_CommandLine = $App.CommandLine
	$App_SilentSwitch = $App.Silent_Switch

	# [System.Windows.Forms.MessageBox]::Show("$App_File")


	$Folder_Name = $App_Path.split("\")[-1]
	$App_Folder = "$Desktop\$Folder_Name"
	$App_Full_Path = "$App_Folder\$App_File"
	
	
	If ( ($App_File -like "*.exe*") -or ($App_File -like "*.msi*") ) {
		If ($App_SilentSwitch -ne "") {
			Start-Process $App_Full_Path -ArgumentList "$App_SilentSwitch" -Wait
		}
		Else {
			Start-Process $App_Full_Path -Wait
		}
	}
	ElseIf ( ($App_File -like "*.ps1*") -or ($App_File -like "*.vbs*") ) {
		& { Invoke-Expression ($App_Full_Path) }
	}
	ElseIf ($App_File -like "*.intunewin") {
		$Config_Folder_Path = "$Desktop\Intunewin_Config_Folder\$Folder_Name"
		New-Item $Config_Folder_Path -Type Directory -Force
		$Intunewin_Content_File = "$Config_Folder_Path\Intunewin_Folder.txt"
		$Intunewin_Command_File = "$Config_Folder_Path\Intunewin_Install_Command.txt"
		
		$App_Full_Path | Out-File $Intunewin_Content_File -Force -NoNewline
		$App_CommandLine | Out-File $Intunewin_Command_File -Force -NoNewline
		C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -sta -WindowStyle Hidden -NoProfile -ExecutionPolicy Unrestricted -File $Sandbox_Root_Path\IntuneWin_Install.ps1 $Intunewin_Content_File $Intunewin_Command_File
	}
	Else {
		Set-Location $App_Folder
		& { Invoke-Expression (Get-Content -Raw $App_File) }
		& { Invoke-Expression ($App_CommandLine) }
	}
}
Read-Host -Prompt "Press Enter to exit"