$Desktop = "C:\Users\WDAGUtilityAccount\Desktop"
$Sandbox_Root_Path = "C:\Run_in_Sandbox"
$App_Bundle_File = "$Sandbox_Root_Path\App_Bundle.sdbapp"
$SDBApp_Root_Path = "C:\SBDApp"
$Get_Apps_to_install = [xml](Get-Content $App_Bundle_File)
$Apps_to_install = $Get_Apps_to_install.Applications.Application
foreach ($App in $Apps_to_install) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

    $App_File = $App.File
    $App_Path = $App.Path
    if ($App_Path) {
        $Folder_Name = $App_Path.split("\")[-1]
        $App_Folder = "$SDBApp_Root_Path\$Folder_Name"
        $App_Full_Path = "$App_Folder\$App_File"
    } else {
        $App_Folder = "$SDBApp_Root_Path"
    }

    $App_CommandLine = $App.CommandLine
    $App_SilentSwitch = $App.Silent_Switch

    if ( ($App_File -like "*.exe*") -or ($App_File -like "*.msi*") ) {
        if ($App_SilentSwitch -ne "") {
            Start-Process $App_Full_Path -ArgumentList "$App_SilentSwitch" -Wait
        } else {
            Start-Process $App_Full_Path -Wait
        }
    } elseif ( ($App_File -like "*.ps1*") -or ($App_File -like "*.vbs*") ) {
        & { Invoke-Expression ($App_Full_Path) }
    } elseif ($App_File -like "*.intunewin") {
        $Config_Folder_Path = "$Desktop\Intunewin_Config_Folder"
        New-Item -Path $Desktop -Name "Intunewin_Config_Folder" -Type Directory -Force
        $Intunewin_Content_File = "$Config_Folder_Path\Intunewin_Folder.txt"
        $Intunewin_Command_File = "$Config_Folder_Path\Intunewin_Install_Command.txt"

        $App_Full_Path | Out-File $Intunewin_Content_File -Force -NoNewline
        $App_CommandLine | Out-File $Intunewin_Command_File -Force -NoNewline
        C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -sta -WindowStyle Normal -NoProfile -ExecutionPolicy Unrestricted -File $Sandbox_Root_Path\IntuneWin_Install.ps1 $Intunewin_Content_File $Intunewin_Command_File
    } else {
        Set-Location $App_Folder
        & { Invoke-Expression (Get-Content -Raw $App_File) }
        & { Invoke-Expression ($App_CommandLine) }
    }
}