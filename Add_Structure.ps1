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
    [Switch]$NoSilent,
    [Switch]$NoCheckpoint
)

$TEMP_Folder = $env:temp
$Log_File = "$TEMP_Folder\RunInSandbox_Install.log"
$Current_Folder = Split-Path $MyInvocation.MyCommand.Path

$Script:Current_User_SID = (Get-ChildItem -Path Registry::\HKEY_USERS | Where-Object { Test-Path -Path "$($_.pspath)\Volatile Environment" } | ForEach-Object { (Get-ItemProperty -Path "$($_.pspath)\Volatile Environment") }).PSParentPath.split("\")[-1]
$HKCU = "Registry::HKEY_USERS\$Current_User_SID"
$HKCU_Classes = "Registry::HKEY_USERS\$Current_User_SID" + "_Classes"
$Windows_Version = (Get-CimInstance -class Win32_OperatingSystem).Caption

if (Test-Path -Path $Log_File) {
    Remove-Item -Path $Log_File
}
New-Item -Path $Log_File -Type file -Force | Out-Null

Function Write-LogMessage([string]$Message, [string]$Message_Type) {
    $MyDate = "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
    Add-Content -Path $Log_File -Value "$MyDate - $Message_Type : $Message"
    Write-Output "$MyDate - $Message_Type : $Message"
}

Function Export-RegConfig {
    param (
        $Reg_Path,
        $Backup_Path
    )

    if (Test-Path -Path "Registry::HKEY_CLASSES_ROOT\$Reg_Path") {
        try {
            reg export "HKEY_CLASSES_ROOT\$Reg_Path" $Backup_Path /y | Out-Null
            Write-LogMessage -Message_Type "SUCCESS" -Message "$Reg_Path has been exported"
        } catch {
            Write-LogMessage -Message_Type "ERROR" -Message "$Reg_Path has not been exported"
        }
    } else {
        Write-LogMessage -Message_Type "INFO" -Message "Can not find registry path: Registry::HKEY_CLASSES_ROOT\$Reg_Path"
    }
    Add-Content -Path $log_file -Value ""
}


Write-LogMessage -Message_Type "INFO" -Message "Starting the configuration of RunInSandbox"

[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
$Run_As_Admin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if ($Run_As_Admin -eq $False) {
    Write-LogMessage -Message_Type "ERROR" -Message "The script has not been launched with admin rights"
    [System.Windows.Forms.MessageBox]::Show("Please run the tool with admin rights :-)")
    break
}

Write-LogMessage -Message_Type "SUCCESS" -Message "The script has been launched with admin rights"

$Is_Sandbox_Installed = (Get-WindowsOptionalFeature -Online | Where-Object { $_.featurename -eq "Containers-DisposableClientVM" }).state
if ($Is_Sandbox_Installed -eq "Disabled") {
    Write-LogMessage -Message_Type "ERROR" -Message "The feature `"Windows Sandbox`" is not installed !!!"
    [System.Windows.Forms.MessageBox]::Show("The feature `"Windows Sandbox`" is not installed !!!")
    break
}

$Current_Folder = Split-Path -Path $MyInvocation.MyCommand.Path
$Sources = $Current_Folder + "\" + "Sources\*"
if (-not (Test-Path -Path $Sources) ) {
    Write-LogMessage -Message_Type "ERROR" -Message "Sources folder is missing"
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [System.Windows.Forms.MessageBox]::Show("It seems you haven´t downloaded all the folder structure.`nThe folder `"Sources`" is missing !!!")
    break
}

Write-LogMessage -Message_Type "SUCCESS" -Message "The sources folder exists"

Add-Content -Path $log_file -Value ""

$Progress_Activity = "Enabling Run in Sandbox context menus"
Write-Progress -Activity $Progress_Activity -PercentComplete 1

$Check_Sources_Files_Count = (Get-ChildItem -Path "$Current_Folder\Sources\Run_in_Sandbox" -Recurse).count

if ($Check_Sources_Files_Count -lt 40) {
    Write-LogMessage -Message_Type "ERROR" -Message "Some contents are missing"
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [System.Windows.Forms.MessageBox]::Show("It seems you haven´t downloaded all the folder structure !!!")
    break
}

$Destination_folder = "$env:ProgramData\Run_in_Sandbox"
try {
    Copy-Item -Path $Sources -Destination $env:ProgramData -Force -Recurse | Out-Null
    Write-LogMessage -Message_Type "SUCCESS" -Message "Sources have been copied in $env:ProgramData\Run_in_Sandbox"
} catch {
    Write-LogMessage -Message_Type "ERROR" -Message "Sources have not been copied in $env:ProgramData\Run_in_Sandbox"
    EXIT
}

$Sources_Unblocked = $False
try {
    Get-ChildItem -Path $Destination_folder -Recurse | Unblock-File
    Write-LogMessage -Message_Type "SUCCESS" -Message "Sources files have been unblocked"
    $Sources_Unblocked = $True
} catch {
    Write-LogMessage -Message_Type "ERROR" -Message "Sources files have not been unblocked"
    break
}

if ($Sources_Unblocked -ne $True) {
    Write-LogMessage -Message_Type "ERROR" -Message "Source files could not be unblocked"
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [System.Windows.Forms.MessageBox]::Show("Source files could not be unblocked")
    break
}

if ($NoSilent) {
    Set-Location -Path "$Current_Folder\Sources\Run_in_Sandbox"
    powershell -NoProfile .\RunInSandbox_Config.ps1
}

$Sandbox_Icon = "$env:ProgramData\Run_in_Sandbox\sandbox.ico"
$Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"
$XML_Config = "$Run_in_Sandbox_Folder\Sandbox_Config.xml"
$Get_XML_Content = [xml](Get-Content $XML_Config)

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

if (-not (Test-Path -Path "$env:ProgramData\Run_in_Sandbox\RunInSandbox.ps1") ) {
    Write-LogMessage -Message_Type "ERROR" -Message "File RunInSandbox.ps1 is missing"
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [System.Windows.Forms.MessageBox]::Show("File RunInSandbox.ps1 is missing !!!")
    break
}

$Backup_Folder = "$Destination_folder\Registry_Backup"
New-Item -Path $Backup_Folder -Type Directory -Force | Out-Null

Write-Progress -Activity $Progress_Activity -PercentComplete 5

Write-LogMessage -Message_Type "INFO" -Message "Exporting registry keys"

Export-RegConfig -Reg_Path "exefile" -Backup_Path "$Backup_Folder\Backup_HKRoot_EXEFile.reg"
Export-RegConfig -Reg_Path "Microsoft.PowerShellScript.1" -Backup_Path "$Backup_Folder\Backup_HKRoot_PowerShellScript.reg"
Export-RegConfig -Reg_Path "VBSFile" -Backup_Path "$Backup_Folder\Backup_HKRoot_VBSFile.reg"
Export-RegConfig -Reg_Path "Msi.Package" -Backup_Path "$Backup_Folder\Backup_HKRoot_Msi.reg"
Export-RegConfig -Reg_Path "CompressedFolder" -Backup_Path "$Backup_Folder\Backup_HKRoot_CompressedFolder.reg"
Export-RegConfig -Reg_Path "WinRAR.ZIP" -Backup_Path "$Backup_Folder\Backup_HKRoot_WinRAR.reg"
Export-RegConfig -Reg_Path "Directory" -Backup_Path "$Backup_Folder\Backup_HKRoot_Directory.reg"

Write-Progress -Activity $Progress_Activity -PercentComplete 10

if (-not $NoCheckpoint) {
    $Checkpoint_Command = 'Checkpoint-Computer -Description "Windows_Sandbox_Context_menus" -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop'
    $ReturnValue = Start-Process powershell -WindowStyle Hidden -ArgumentList $Checkpoint_Command -Wait -PassThru
    if ($ReturnValue.ExitCode -eq 0) {
        Write-LogMessage -Message_Type "SUCCESS" -Message "Creation of restore point `"Add Windows Sandbox Context menus`""
    } else {
        Write-LogMessage -Message_Type "ERROR" -Message "Creation of restore point `"Add Windows Sandbox Context menus`""
    }
}

Function Add-RegKey {
    param (
        $Reg_Path = "Registry::HKEY_CLASSES_ROOT",
        $Sub_Reg_Path,
        $Type,
        $Entry_Name = $Type,
        $Info_Type = $Type,
        $Key_Label = "Run $Entry_Name in Sandbox"
    )

    try {
        $Base_Registry_Key = "$Reg_Path\$Sub_Reg_Path"
        $Shell_Registry_Key = "$Base_Registry_Key\Shell"
        $Key_Label_Path = "$Shell_Registry_Key\$Key_Label"
        $Command_Path = "$Key_Label_Path\Command"
        $Command_for = "C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -WindowStyle Hidden -NoProfile -ExecutionPolicy Unrestricted -sta -File C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -Type $Type -ScriptPath `"%V`""
        if (-not (Test-Path -Path $Base_Registry_Key) ) {
            New-Item -Path $Base_Registry_Key -ErrorAction Stop | Out-Null
        }
        if (-not (Test-Path -Path $Shell_Registry_Key) ) {
            New-Item -Path $Shell_Registry_Key -ErrorAction Stop | Out-Null
        }
        if (Test-Path -Path $Key_Label_Path) {
            Write-LogMessage -Message_Type "SUCCESS" -Message "Context menu for $Type has already been added"
            return
        }

        New-Item -Path $Key_Label_Path -ErrorAction Stop | Out-Null
        New-Item -Path $Command_Path -ErrorAction Stop | Out-Null
        # Add Sandbox Icons
        New-ItemProperty -Path $Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon -ErrorAction Stop | Out-Null
        # Set the command path
        Set-Item -Path $Command_Path -Value $Command_for -Force -ErrorAction Stop | Out-Null
        Write-LogMessage -Message_Type "SUCCESS" -Message "Context menu for `"$Info_Type`" has been added"
    } catch {
        Write-LogMessage -Message_Type "ERROR" -Message "Context menu for $Type couldn´t be added"
    }

}

Write-Progress -Activity $Progress_Activity -PercentComplete 20

Write-LogMessage -Message_Type "INFO" -Message "Adding context menu"
Write-LogMessage -Message_Type "INFO" -Message "OS version is: $Windows_Version"

if ($Add_PS1 -eq $True) {
    $PS1_Main_Menu = "Run PS1 in Sandbox"

    if ($Windows_Version -like "*Windows 10*") {
        Write-LogMessage -Message_Type "INFO" -Message "Running on Windows 10"
        $PS1_Shell_Registry_Key = "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\.ps1\Shell"
        $Main_Menu_Path = "$PS1_Shell_Registry_Key\$PS1_Main_Menu"
        New-Item -Path $PS1_Shell_Registry_Key -Name $PS1_Main_Menu -Force | Out-Null
        New-ItemProperty -Path $Main_Menu_Path -Name "subcommands" -PropertyType String | Out-Null
        New-Item -Path $Main_Menu_Path -Name "Shell" -Force | Out-Null

        Add-RegKey -Sub_Reg_Path "SystemFileAssociations\.ps1\Shell\Run PS1 in Sandbox" -Type "PS1Basic" -Entry_Name "PS1 as user"
        Add-RegKey -Sub_Reg_Path "SystemFileAssociations\.ps1\Shell\Run PS1 in Sandbox" -Type "PS1System" -Entry_Name "PS1 as system"
        Add-RegKey -Sub_Reg_Path "SystemFileAssociations\.ps1\Shell\Run PS1 in Sandbox" -Type "PS1Params" -Entry_Name "PS1 with Parameters"
        New-ItemProperty -Path "$Main_Menu_Path" -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
    }
    if ($Windows_Version -like "*Windows 11*") {
        Write-LogMessage -Message_Type "INFO" -Message "Running on Windows 11"

        $Default_PS1_HKCU = "$HKCU_Classes\.ps1"
        if (Test-Path -Path $HKCU_Classes) {
            $rOpenWithProgids_Key = "$Default_PS1_HKCU\rOpenWithProgids"
            if (Test-Path -Path $rOpenWithProgids_Key) {
                $Get_OpenWithProgids_Default_Value = (Get-Item -Path $rOpenWithProgids_Key).Property
                ForEach ($Prop in $Get_OpenWithProgids_Default_Value) {
                    $PS1_Shell_Registry_Key = "$HKCU_Classes\$Prop\Shell"
                    $Main_Menu_Path = "$PS1_Shell_Registry_Key\$PS1_Main_Menu"
                    New-Item -Path $PS1_Shell_Registry_Key -Name $PS1_Main_Menu -Force | Out-Null
                    New-ItemProperty -Path $Main_Menu_Path -Name "subcommands" -PropertyType String | Out-Null
                    New-Item -Path $Main_Menu_Path -Name "Shell" -Force | Out-Null

                    Add-RegKey -Reg_Path "$PS1_Shell_Registry_Key" -Sub_Reg_Path "$PS1_Main_Menu" -Type "PS1Basic" -Entry_Name "PS1 as user"
                    Add-RegKey -Reg_Path "$PS1_Shell_Registry_Key" -Sub_Reg_Path "$PS1_Main_Menu" -Type "PS1System" -Entry_Name "PS1 as system"
                    Add-RegKey -Reg_Path "$PS1_Shell_Registry_Key" -Sub_Reg_Path "$PS1_Main_Menu" -Type "PS1Params" -Entry_Name "PS1 with Parameters"
                }
            }
            $OpenWithProgids_Key = "$Default_PS1_HKCU\OpenWithProgids"
            if (Test-Path -Path $OpenWithProgids_Key) {
                $Get_OpenWithProgids_Default_Value = (Get-Item -Path $OpenWithProgids_Key).Property
                ForEach ($Prop in $Get_OpenWithProgids_Default_Value) {
                    $PS1_Shell_Registry_Key = "$HKCU_Classes\$Prop\Shell"
                    $Main_Menu_Path = "$PS1_Shell_Registry_Key\$PS1_Main_Menu"
                    New-Item -Path $PS1_Shell_Registry_Key -Name $PS1_Main_Menu -Force | Out-Null
                    New-ItemProperty -Path $Main_Menu_Path -Name "subcommands" -PropertyType String | Out-Null

                    Add-RegKey -Reg_Path "$PS1_Shell_Registry_Key" -Sub_Reg_Path "$PS1_Main_Menu" -Type "PS1Basic" -Entry_Name "PS1 as user"
                    Add-RegKey -Reg_Path "$PS1_Shell_Registry_Key" -Sub_Reg_Path "$PS1_Main_Menu" -Type "PS1System" -Entry_Name "PS1 as system"
                    Add-RegKey -Reg_Path "$PS1_Shell_Registry_Key" -Sub_Reg_Path "$PS1_Main_Menu" -Type "PS1Params" -Entry_Name "PS1 with Parameters"
                }
            }

            # ADDING CONTEXT MENU DEPENDING OF THE USERCHOICE
            # The userchoice for PS1 is located in: HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.ps1\UserChoice
            $PS1_UserChoice = "$HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.ps1\UserChoice"
            $Get_UserChoice = (Get-ItemProperty -Path $PS1_UserChoice).ProgID

            $HKCR_UserChoice_Key = "Registry::HKEY_CLASSES_ROOT\$Get_UserChoice"
            $PS1_Shell_Registry_Key = "$HKCR_UserChoice_Key\Shell"
            if (Test-Path -Path $PS1_Shell_Registry_Key) {
                $Main_Menu_Path = "$PS1_Shell_Registry_Key\$PS1_Main_Menu"
                New-Item -Path $PS1_Shell_Registry_Key -Name $PS1_Main_Menu -Force | Out-Null
                New-ItemProperty -Path $Main_Menu_Path -Name "subcommands" -PropertyType String | Out-Null

                Add-RegKey -Reg_Path "$PS1_Shell_Registry_Key" -Sub_Reg_Path "$PS1_Main_Menu" -Type "PS1Basic" -Entry_Name "PS1 as user"
                Add-RegKey -Reg_Path "$PS1_Shell_Registry_Key" -Sub_Reg_Path "$PS1_Main_Menu" -Type "PS1System" -Entry_Name "PS1 as system"
                Add-RegKey -Reg_Path "$PS1_Shell_Registry_Key" -Sub_Reg_Path "$PS1_Main_Menu" -Type "PS1Params" -Entry_Name "PS1 with Parameters"
            }
        }
        New-ItemProperty -Path "$Main_Menu_Path" -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
    }
}
Write-Progress -Activity $Progress_Activity -PercentComplete 25

if ($Add_Intunewin -eq $True) {
    Add-RegKey -Sub_Reg_Path ".intunewin" -Type "Intunewin"
}
Write-Progress -Activity $Progress_Activity -PercentComplete 30


if ($Add_Reg -eq $True) {
    Add-RegKey -Sub_Reg_Path "regfile" -Type "REG" -Key_Label "Test reg file in Sandbox"
}
Write-Progress -Activity $Progress_Activity -PercentComplete 35

if ($Add_ISO -eq $True) {
    Add-RegKey -Sub_Reg_Path "Windows.IsoFile" -Type "ISO" -Key_Label "Extract ISO file in Sandbox"
    Add-RegKey -Reg_Path "$HKCU_Classes" -Sub_Reg_Path ".iso" -Type "ISO" -Key_Label "Extract ISO file in Sandbox"
}
Write-Progress -Activity $Progress_Activity -PercentComplete 40

if ($Add_PPKG -eq $True) {
    Add-RegKey -Sub_Reg_Path "Microsoft.ProvTool.Provisioning.1" -Type "PPKG"
}
Write-Progress -Activity $Progress_Activity -PercentComplete 45

if ($Add_HTML -eq $True) {
    Add-RegKey -Sub_Reg_Path "MSEdgeHTM" -Type "HTML" -Key_Label "Run this web link in Sandbox"
    Add-RegKey -Sub_Reg_Path "ChromeHTML" -Type "HTML" -Key_Label "Run this web link in Sandbox"
    Add-RegKey -Sub_Reg_Path "IE.AssocFile.HTM" -Type "HTML" -Key_Label "Run this web link in Sandbox"
    Add-RegKey -Sub_Reg_Path "IE.AssocFile.URL" -Type "HTML" -Key_Label "Run this URL in Sandbox"
}
Write-Progress -Activity $Progress_Activity -PercentComplete 50

if ($Add_MultipleApp -eq $True) {
    Add-RegKey -Sub_Reg_Path ".sdbapp" -Type "SDBApp" -Entry_Name "application bundle"
}
Write-Progress -Activity $Progress_Activity -PercentComplete 55

if ($Add_VBS -eq $True) {
    $VBS_Main_Menu = "Run VBS in Sandbox"

    $VBS_Shell_Registry_Key = "Registry::HKEY_CLASSES_ROOT\VBSFile\Shell"
    $Main_Menu_Path = "$VBS_Shell_Registry_Key\$VBS_Main_Menu"
    New-Item -Path $VBS_Shell_Registry_Key -Name $VBS_Main_Menu -Force | Out-Null
    New-ItemProperty -Path $Main_Menu_Path -Name "subcommands" -PropertyType String | Out-Null
    New-Item -Path $Main_Menu_Path -Name "Shell" -Force | Out-Null

    Add-RegKey -Sub_Reg_Path "VBSFile\Shell\$VBS_Main_Menu" -Type "VBSBasic" -Entry_Name "VBS"
    Add-RegKey -Sub_Reg_Path "VBSFile\Shell\$VBS_Main_Menu" -Type "VBSParams" -Entry_Name "VBS with Parameters"

    New-ItemProperty -Path "$Main_Menu_Path" -Name "icon" -PropertyType String -Value $Sandbox_Icon | Out-Null
}
Write-Progress -Activity $Progress_Activity -PercentComplete 60


if ($Add_EXE -eq $True) {
    Add-RegKey -Sub_Reg_Path "exefile" -Type "EXE"
}
Write-Progress -Activity $Progress_Activity -PercentComplete 65

if ($Add_MSI -eq $True) {
    Add-RegKey -Sub_Reg_Path "Msi.Package" -Type "MSI"
}
Write-Progress -Activity $Progress_Activity -PercentComplete 70

if ($Add_ZIP -eq $True) {
    # Run on ZIP
    Add-RegKey -Sub_Reg_Path "CompressedFolder" -Type "ZIP" -Key_Label "Extract ZIP in Sandbox"

    # Run on ZIP if WinRAR is installed
    if (Test-Path -Path "Registry::HKEY_CLASSES_ROOT\WinRAR.ZIP") {
        Add-RegKey -Sub_Reg_Path "WinRAR.ZIP" -Type "ZIP" -Key_Label "Extract ZIP (WinRAR) in Sandbox"
    }

    # Run on 7z
    if (Test-Path -Path "Registry::HKEY_CLASSES_ROOT\Applications\7zFM.exe") {
        Add-RegKey -Sub_Reg_Path "Applications\7zFM.exe" -Type "7z" -Info_Type "7z" -Entry_Name "ZIP" -Key_Label "Extract 7z file in Sandbox"
    }
    if (Test-Path -Path "Registry::HKEY_CLASSES_ROOT\7-Zip.7z") {
        Add-RegKey -Sub_Reg_Path "7-Zip.7z" -Type "7z" -Info_Type "7z" -Entry_Name "ZIP" -Key_Label "Extract 7z file in Sandbox"
    }
    if (Test-Path -Path "Registry::HKEY_CLASSES_ROOT\7-Zip.rar") {
        Add-RegKey -Sub_Reg_Path "7-Zip.rar" -Type "7z" -Info_Type "7z" -Entry_Name "ZIP" -Key_Label "Extract RAR file in Sandbox"
    }
}
Write-Progress -Activity $Progress_Activity -PercentComplete 75

if ($Add_MSIX -eq $True) {
    $MSIX_Shell_Registry_Key = "Registry::HKEY_CLASSES_ROOT\.msix\OpenWithProgids"
    if (Test-Path -Path $MSIX_Shell_Registry_Key) {
        $Get_Default_Value = (Get-Item -Path $MSIX_Shell_Registry_Key).Property
        Add-RegKey -Sub_Reg_Path "$Get_Default_Value" -Type "MSIX"
    }
    if (Test-Path -Path $HKCU_Classes) {
        $Default_MSIX_HKCU = "$HKCU_Classes\.msix"
        $Get_Default_Value = (Get-Item -Path "$Default_MSIX_HKCU\OpenWithProgids").Property
        Add-RegKey -Reg_Path $HKCU_Classes -Sub_Reg_Path "$Get_Default_Value" -Type "MSIX"
    }
}
Write-Progress -Activity $Progress_Activity -PercentComplete 80

if ($Add_Folder -eq $True) {
    Add-RegKey -Sub_Reg_Path "Directory\Background" -Type "Folder_Inside" -Entry_Name "this folder" -Key_Label "Share this folder in a Sandbox"
    Add-RegKey -Sub_Reg_Path "Directory" -Type "Folder_On" -Entry_Name "this folder" -Key_Label "Share this folder in a Sandbox"
}
Write-Progress -Activity $Progress_Activity -PercentComplete 85

if ($Add_CMD -eq $True) {
    Add-RegKey -Sub_Reg_Path "cmdfile" -Type "CMD"
    Add-RegKey -Sub_Reg_Path "batfile" -Type "BAT"
}
Write-Progress -Activity $Progress_Activity -PercentComplete 90

if ($Add_PDF -eq $True) {

    if (-not (Test-Path -Path "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\.pdf") ) {
        New-Item -Path "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\.pdf" -ErrorAction Stop | Out-Null
    }
    Add-RegKey -Sub_Reg_Path "SystemFileAssociations\.pdf" -Type "PDF" -Key_Label "Open PDF in Sandbox"
}
Write-Progress -Activity $Progress_Activity -PercentComplete 100

Copy-Item -Path $Log_File -Destination $Destination_folder -Force


