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

param (
    [Parameter(Mandatory=$true)] [String]$Type,
    [Parameter(Mandatory=$true)] [String]$ScriptPath
)

$special_char_array = 'é', 'è', 'à', 'â', 'ê', 'û', 'î', 'ä', 'ë', 'ü', 'ï', 'ö', 'ù', 'ò', '~', '!', '@', '#', '$', '%', '^', '&', '+', '=', '}', '{', '|', '<', '>', ';'
foreach ($char in $special_char_array) {
    if ($ScriptPath -like "*$char*") {
        [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        $message = "There is a special character in the path of the file (`'" + $char + "`').`nWindows Sandbox does not support this!"
        [System.Windows.Forms.MessageBox]::Show($message, "Issue with your file", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        EXIT
    }
}

$ScriptPath = $ScriptPath.replace('"', '')
$ScriptPath = $ScriptPath.Trim();
$ScriptPath = [WildcardPattern]::Escape($ScriptPath)

if ( ($Type -eq "Folder_Inside") -or ($Type -eq "Folder_On") ) {
    $DirectoryName = (Get-Item $ScriptPath).fullname
} else {
    $FolderPath = Split-Path (Split-Path "$ScriptPath" -Parent) -Leaf
    $DirectoryName = (Get-Item $ScriptPath).DirectoryName
    $FileName = (Get-Item $ScriptPath).BaseName
    $Full_FileName = (Get-Item $ScriptPath).Name
}

$Sandbox_Desktop_Path = "C:\Users\WDAGUtilityAccount\Desktop"
$Sandbox_Shared_Path = "$Sandbox_Desktop_Path\$FolderPath"

$Sandbox_Root_Path = "C:\Run_in_Sandbox"
$Full_Startup_Path = "$Sandbox_Shared_Path\$Full_FileName"
$Full_Startup_Path_Quoted = """$Full_Startup_Path"""

$Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"

$xml = "$Run_in_Sandbox_Folder\Sandbox_Config.xml"
$my_xml = [xml](Get-Content $xml)
$Sandbox_VGpu = $my_xml.Configuration.VGpu
$Sandbox_Networking = $my_xml.Configuration.Networking
$Sandbox_ReadOnlyAccess = $my_xml.Configuration.ReadOnlyAccess
$Sandbox_WSB_Location = $my_xml.Configuration.WSB_Location
$Sandbox_AudioInput = $my_xml.Configuration.AudioInput
$Sandbox_VideoInput = $my_xml.Configuration.VideoInput
$Sandbox_ProtectedClient = $my_xml.Configuration.ProtectedClient
$Sandbox_PrinterRedirection = $my_xml.Configuration.PrinterRedirection
$Sandbox_ClipboardRedirection = $my_xml.Configuration.ClipboardRedirection
$Sandbox_MemoryInMB = $my_xml.Configuration.MemoryInMB
$WSB_Cleanup = $my_xml.Configuration.WSB_Cleanup
$Hide_Powershell = $my_xml.Configuration.Hide_Powershell

[System.Collections.ArrayList]$PowershellParameters = @(
    '-sta'
    '-WindowStyle'
    'Hidden'
    '-NoProfile'
    '-ExecutionPolicy'
    'Unrestricted'
)

if ($Hide_Powershell -eq "False") {
    $PowershellParameters[[array]::IndexOf($PowershellParameters, "Hidden")] = "Normal"
}

$PSRun_File = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe $PowershellParameters -File"
$PSRun_Command = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe $PowershellParameters -Command"

if ($Sandbox_WSB_Location -eq "Default") {
    $Sandbox_File_Path = "$env:temp\$FileName.wsb"
} else {
    $Sandbox_File_Path = "$Sandbox_WSB_Location\$FileName.wsb"
}

if (Test-Path $Sandbox_File_Path) {
    Remove-Item $Sandbox_File_Path
}


function New-WSB {
    param (
        [String]$Command_to_Run
    )

    New-Item $Sandbox_File_Path -type file -Force | Out-Null
    Add-Content -LiteralPath $Sandbox_File_Path -Value "<Configuration>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "    <VGpu>$Sandbox_VGpu</VGpu>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "    <Networking>$Sandbox_Networking</Networking>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "    <AudioInput>$Sandbox_AudioInput</AudioInput>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "    <VideoInput>$Sandbox_VideoInput</VideoInput>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "    <ProtectedClient>$Sandbox_ProtectedClient</ProtectedClient>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "    <PrinterRedirection>$Sandbox_PrinterRedirection</PrinterRedirection>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "    <ClipboardRedirection>$Sandbox_ClipboardRedirection</ClipboardRedirection>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "    <MemoryInMB>$Sandbox_MemoryInMB</MemoryInMB>"

    Add-Content $Sandbox_File_Path "    <MappedFolders>"
    if ( ($Type -eq "Intunewin") -or ($Type -eq "ISO") -or ($Type -eq "7z")  -or ($Type -eq "PS1System") -or ($Type -eq "SDBApp") ) {
        Add-Content -LiteralPath $Sandbox_File_Path -Value "        <MappedFolder>"
        Add-Content -LiteralPath $Sandbox_File_Path -Value "            <HostFolder>C:\ProgramData\Run_in_Sandbox</HostFolder>"
        Add-Content -LiteralPath $Sandbox_File_Path -Value "            <SandboxFolder>C:\Run_in_Sandbox</SandboxFolder>"
        Add-Content -LiteralPath $Sandbox_File_Path -Value "            <ReadOnly>$Sandbox_ReadOnlyAccess</ReadOnly>"
        Add-Content -LiteralPath $Sandbox_File_Path -Value "        </MappedFolder>"
    }

    if ($Type -eq "SDBApp") {
        $SDB_Full_Path = $ScriptPath
        Copy-Item $ScriptPath $Run_in_Sandbox_Folder -Force
        $Get_Apps_to_install = [xml](Get-Content $SDB_Full_Path)
        $Apps_to_install_path = $Get_Apps_to_install.Applications.Application.Path | Select-Object -Unique

        ForEach ($App_Path in $Apps_to_install_path) {
            Add-Content -LiteralPath $Sandbox_File_Path -Value "        <MappedFolder>"
            Add-Content -LiteralPath $Sandbox_File_Path -Value "            <HostFolder>$App_Path</HostFolder>"
            Add-Content -LiteralPath $Sandbox_File_Path -Value "            <SandboxFolder>C:\SBDApp</SandboxFolder>"
            Add-Content -LiteralPath $Sandbox_File_Path -Value "            <ReadOnly>$Sandbox_ReadOnlyAccess</ReadOnly>"
            Add-Content -LiteralPath $Sandbox_File_Path -Value "        </MappedFolder>"
        }
    } else {
        Add-Content -LiteralPath $Sandbox_File_Path -Value "        <MappedFolder>"
        Add-Content -LiteralPath $Sandbox_File_Path -Value "            <HostFolder>$DirectoryName</HostFolder>"
        if ($Type -eq "IntuneWin") { Add-Content -LiteralPath $Sandbox_File_Path -Value "           <SandboxFolder>C:\IntuneWin</SandboxFolder>" }
        Add-Content -LiteralPath $Sandbox_File_Path -Value "            <ReadOnly>$Sandbox_ReadOnlyAccess</ReadOnly>"
        Add-Content -LiteralPath $Sandbox_File_Path -Value "        </MappedFolder>"
    }
    Add-Content -LiteralPath $Sandbox_File_Path -Value "    </MappedFolders>"
}

switch ($Type) {
    "7Z" {
        $Script:Startup_Command = "$Sandbox_Root_Path\7z\7z.exe" + " " + "x" + " " + "$Full_Startup_Path_Quoted" + " " + "-y" + " " + "-o" + "C:\Users\WDAGUtilityAccount\Desktop\Extracted_File"
        New-WSB -Command_to_Run $Startup_Command
    }
    "CMD" {
        $Script:Startup_Command = $PSRun_Command + " " + "Start-Process $Full_Startup_Path_Quoted"
        New-WSB -Command_to_Run $Startup_Command
    }
    "EXE" {
        $Full_Startup_Path_Quoted = $Full_Startup_Path_Quoted.Replace('"', "")

        [System.Reflection.Assembly]::LoadWithPartialName('presentationframework') | Out-Null
        [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.dll") | Out-Null
        [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.IconPacks.dll") | Out-Null
        Function LoadXml ($global:file2) {
            $XamlLoader = (New-Object System.Xml.XmlDocument)
            $XamlLoader.Load($file2)
            return $XamlLoader
        }

        $XamlMainWindow = LoadXml("$Run_in_Sandbox_Folder\RunInSandbox_EXE.xaml")
        $Reader = (New-Object System.Xml.XmlNodeReader $XamlMainWindow)
        $Form_EXE = [Windows.Markup.XamlReader]::Load($Reader)
        $EXE_Command_File = "$Run_in_Sandbox_Folder\EXE_Command_File.txt"

        $switches_for_exe = $Form_EXE.findname("switches_for_exe")
        $add_switches = $Form_EXE.findname("add_switches")

        $add_switches.Add_Click({
                $Script:Switches_EXE = $switches_for_exe.Text.ToString()
                $Script:Startup_Command = $Full_Startup_Path_Quoted + " " + $Switches_EXE
                $Startup_Command | Out-File $EXE_Command_File -Force -NoNewline
                $Form_EXE.close()
            })

        $Form_EXE.Add_Closing({
                $Script:Switches_EXE = $switches_for_exe.Text.ToString()
                $Script:Startup_Command = $Full_Startup_Path_Quoted + " " + $Switches_EXE
                $Startup_Command | Out-File $EXE_Command_File -Force -NoNewline
            })

        $Form_EXE.ShowDialog() | Out-Null

        $EXE_Installer = "$Sandbox_Root_Path\EXE_Install.ps1"
        $Script:Startup_Command = $PSRun_File + " " + "$EXE_Installer"
        New-WSB -Command_to_Run $Startup_Command
    }
    "Folder_On" {
        New-WSB
    }
    "Folder_Inside" {
        New-WSB
    }
    "HTML" {
        $Script:Startup_Command = $PSRun_Command + " " + "`"Invoke-Item -Path `'$Full_Startup_Path_Quoted`'`""
        New-WSB -Command_to_Run $Startup_Command
    }
    "URL" {
        $Script:Startup_Command = $PSRun_Command + " " + "Start-Process $Sandbox_Root_Path"
        New-WSB -Command_to_Run $Startup_Command
    }
    "Intunewin" {
        $Intunewin_Folder = "$Sandbox_Shared_Path\$FileName.intunewin"
        $Intunewin_Content_File = "$Run_in_Sandbox_Folder\Intunewin_Folder.txt"
        $Intunewin_Command_File = "$Run_in_Sandbox_Folder\Intunewin_Install_Command.txt"
        $Intunewin_Folder | Out-File $Intunewin_Content_File -Force -NoNewline

        $Full_Startup_Path_Quoted = $Full_Startup_Path_Quoted.Replace('"', "")

        [System.Reflection.Assembly]::LoadWithPartialName('presentationframework')  | Out-Null
        [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.dll") | Out-Null
        [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.IconPacks.dll") | Out-Null
        function LoadXml ($global:file1) {
            $XamlLoader = (New-Object System.Xml.XmlDocument)
            $XamlLoader.Load($file1)
            return $XamlLoader
        }

        $XamlMainWindow = LoadXml("$Run_in_Sandbox_Folder\RunInSandbox_Intunewin.xaml")
        $Reader = (New-Object System.Xml.XmlNodeReader $XamlMainWindow)
        $Form_PS1 = [Windows.Markup.XamlReader]::Load($Reader)

        $install_command_intunewin = $Form_PS1.findname("install_command_intunewin")
        $add_install_command = $Form_PS1.findname("add_install_command")

        $add_install_command.add_click({
                $Script:install_command = $install_command_intunewin.Text.ToString()
                $install_command | Out-File $Intunewin_Command_File
                $Form_PS1.close()
            })

        $Form_PS1.Add_Closing({
                $Script:install_command = $install_command_intunewin.Text.ToString()
                $install_command | Out-File $Intunewin_Command_File -Force -NoNewline
                $Form_PS1.close()
            })

        $Form_PS1.ShowDialog() | Out-Null

        $Intunewin_Installer = "$Sandbox_Root_Path\IntuneWin_Install.ps1"
        $Script:Startup_Command = $PSRun_File + " " + "$Intunewin_Installer"
        New-WSB -Command_to_Run $Startup_Command
    }
    "ISO" {
        $Script:Startup_Command = "$Sandbox_Root_Path\7z\7z.exe" + " " + "x" + " " + "$Full_Startup_Path_Quoted" + " " + "-y" + " " + "-o" + "C:\Users\WDAGUtilityAccount\Desktop\Extracted_ISO"
        New-WSB -Command_to_Run $Startup_Command
    }
    "MSI" {
        $Full_Startup_Path_Quoted = $Full_Startup_Path_Quoted.Replace('"', "")

        [System.Reflection.Assembly]::LoadWithPartialName('presentationframework')              | Out-Null
        [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.dll") | Out-Null
        [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.IconPacks.dll")      | Out-Null
        function LoadXml ($global:file2) {
            $XamlLoader = (New-Object System.Xml.XmlDocument)
            $XamlLoader.Load($file2)
            return $XamlLoader
        }

        $XamlMainWindow = LoadXml("$Run_in_Sandbox_Folder\RunInSandbox_EXE.xaml")
        $Reader = (New-Object System.Xml.XmlNodeReader $XamlMainWindow)
        $Form_MSI = [Windows.Markup.XamlReader]::Load($Reader)

        $switches_for_exe = $Form_MSI.findname("switches_for_exe")
        $add_switches = $Form_MSI.findname("add_switches")

        $add_switches.Add_Click({
                $Script:Switches_MSI = $switches_for_exe.Text.ToString()
                $Script:Startup_Command = "msiexec /i `"$Full_Startup_Path_Quoted`" " + $Switches_MSI
                $Form_MSI.close()
            })

        $Form_MSI.Add_Closing({
                $Script:Switches_MSI = $switches_for_exe.Text.ToString()
                $Script:Startup_Command = "msiexec /i `"$Full_Startup_Path_Quoted`" " + $Switches_MSI
            })

        $Form_MSI.ShowDialog() | Out-Null

        New-WSB -Command_to_Run $Startup_Command
    }
    "MSIX" {
        $Script:Startup_Command = $PSRun_Command + " " + "Add-AppPackage -Path $Full_Startup_Path_Quoted"
        New-WSB -Command_to_Run $Startup_Command
    }
    "PDF" {
        $Full_Startup_Path_Quoted = $Full_Startup_Path_Quoted.Replace('"', '')
        $Script:Startup_Command = $PSRun_Command + " " + "`"Invoke-Item -Path `'$Full_Startup_Path_Quoted`'`""
        New-WSB -Command_to_Run $Startup_Command
    }
    "PPKG" {
        $Script:Startup_Command = $PSRun_Command + " " + "Install-ProvisioningPackage $Full_Startup_Path_Quoted -forceinstall -quietinstall"
        New-WSB -Command_to_Run $Startup_Command
    }
    "PS1Basic" {
        $Script:Startup_Command = $PSRun_File + " " + "$Full_Startup_Path_Quoted"
        New-WSB -Command_to_Run $Startup_Command
    }
    "PS1System" {
        $Script:Startup_Command = "C:\Users\WDAGUtilityAccount\Desktop\Run_in_Sandbox\PsExec.exe -accepteula -i -d -s powershell -executionpolicy bypass -file $Full_Startup_Path_Quoted"
        New-WSB -Command_to_Run $Startup_Command
    }
    "PS1Params" {
        $Full_Startup_Path_Quoted = $Full_Startup_Path_Quoted.Replace('"', "")

        [System.Reflection.Assembly]::LoadWithPartialName('presentationframework')  | Out-Null
        [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.dll") | Out-Null
        [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.IconPacks.dll") | Out-Null
        function LoadXml ($global:file1) {
            $XamlLoader = (New-Object System.Xml.XmlDocument)
            $XamlLoader.Load($file1)
            return $XamlLoader
        }

        $XamlMainWindow = LoadXml("$Run_in_Sandbox_Folder\RunInSandbox_Params.xaml")
        $Reader = (New-Object System.Xml.XmlNodeReader $XamlMainWindow)
        $Form_PS1 = [Windows.Markup.XamlReader]::Load($Reader)

        $parameters_to_add = $Form_PS1.findname("parameters_to_add")
        $add_parameters = $Form_PS1.findname("add_parameters")

        $add_parameters.add_click({
                $Script:Paramaters = $parameters_to_add.Text.ToString()
                $Script:Startup_Command = $PSRun_File + " " + "$Full_Startup_Path_Quoted" + " " + "$Paramaters"
                $Form_PS1.close()
            })

        $Form_PS1.Add_Closing({
                $Script:Paramaters = $parameters_to_add.Text.ToString()
                $Script:Startup_Command = $PSRun_File + " " + "$Full_Startup_Path_Quoted" + " " + "$Paramaters"
            })

        $Form_PS1.ShowDialog() | Out-Null

        New-WSB -Command_to_Run $Startup_Command
    }
    "REG" {
        $Script:Startup_Command = "REG IMPORT $Full_Startup_Path_Quoted"
        New-WSB -Command_to_Run $Startup_Command
    }
    "SDBApp" {
        $AppBundle_Installer = "$Sandbox_Root_Path\AppBundle_Install.ps1"
        $Script:Startup_Command = $PSRun_File + " " + "$AppBundle_Installer"
        New-WSB -Command_to_Run $Startup_Command
    }
    "VBSBasic" {
        $Script:Startup_Command = "wscript.exe $Full_Startup_Path_Quoted"
        New-WSB -Command_to_Run $Startup_Command
    }
    "VBSParams" {
        $Full_Startup_Path_Quoted = $Full_Startup_Path_Quoted.Replace('"', '')

        [System.Reflection.Assembly]::LoadWithPartialName('presentationframework') | Out-Null
        [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.dll") | Out-Null
        [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.IconPacks.dll") | Out-Null
        function LoadXml ($Script:file1) {
            $XamlLoader = (New-Object System.Xml.XmlDocument)
            $XamlLoader.Load($file1)
            return $XamlLoader
        }

        $XamlMainWindow = LoadXml("$Run_in_Sandbox_Folder\RunInSandbox_Params.xaml")
        $Reader = (New-Object System.Xml.XmlNodeReader $XamlMainWindow)
        $Form_VBS = [Windows.Markup.XamlReader]::Load($Reader)

        $parameters_to_add = $Form_VBS.findname("parameters_to_add")
        $add_parameters = $Form_VBS.findname("add_parameters")

        $add_parameters.add_click({
                $Script:Paramaters = $parameters_to_add.Text.ToString()
                $Script:Startup_Command = "wscript.exe $Full_Startup_Path_Quoted $Paramaters"
                $Form_VBS.close()
            })

        $Form_VBS.Add_Closing({
                $Script:Paramaters = $parameters_to_add.Text.ToString()
                $Script:Startup_Command = "wscript.exe $Full_Startup_Path_Quoted $Paramaters"
            })

        $Form_VBS.ShowDialog() | Out-Null

        New-WSB -Command_to_Run $Startup_Command
    }
    "ZIP" {
        $Script:Startup_Command = $PSRun_Command + " " + "`"Expand-Archive -LiteralPath '$Full_Startup_Path' -DestinationPath '$Sandbox_Desktop_Path\ZIP_extracted'`""
        New-WSB -Command_to_Run $Startup_Command
    }
}

Start-Process $Sandbox_File_Path -Wait

if ($WSB_Cleanup -eq $True) {
    Remove-Item -LiteralPath $Sandbox_File_Path -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $Intunewin_Command_File -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $Intunewin_Content_File -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $EXE_Command_File -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath "$Run_in_Sandbox_Folder\App_Bundle.sdbapp" -Force -ErrorAction SilentlyContinue
}