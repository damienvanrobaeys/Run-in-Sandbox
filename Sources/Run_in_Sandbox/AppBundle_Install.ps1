$Desktop = "C:\Users\WDAGUtilityAccount\Desktop"
$Sandbox_Folder = "$Desktop\Run_in_Sandbox"
$App_Bundle_File = "$Sandbox_Folder\App_Bundle.sdbapp"
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

	If ($App_CommandLine -ne "") {
		Set-Location $App_Path
		& { Invoke-Expression (Get-Content -Raw $file) }
		& { Invoke-Expression ($App_CommandLine) }
	}
	Else {
		# set-location $App_Path
		# & $App_Full_Path -wait
		If ( ($App_File -like "*.exe*") -or ($App_File -like "*.msi*") ) {
			If ($App_SilentSwitch -ne "") {
				Start-Process $App_Full_Path -ArgumentList "$App_SilentSwitch" -Wait
			}
			Else {
				Start-Process $App_Full_Path -Wait
			}
		}
		If ( ($App_File -like "*.ps1*") -or ($App_File -like "*.vbs*") ) {
			& { Invoke-Expression ($App_Full_Path) }
	}
}