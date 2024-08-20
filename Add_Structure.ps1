param (
    [Switch]$NoSilent,
    [Switch]$NoCheckpoint
)

$Current_Folder = $PSScriptRoot

Unblock-File -Path $Current_Folder\CommonFunctions.ps1
. "$Current_Folder\CommonFunctions.ps1"


if (Test-Path -Path $Log_File) {
    Remove-Item -Path $Log_File
}
New-Item -Path $Log_File -Type file -Force | Out-Null


Write-LogMessage -Message_Type "INFO" -Message "Starting the configuration of RunInSandbox"

CheckFor-Admin

CheckFor-Sandbox

CheckFor-Sources


$Progress_Activity = "Enabling Run in Sandbox context menus"
Write-Progress -Activity $Progress_Activity -PercentComplete 1


Copy-Sources

Unblock-Sources

if ($NoSilent) {
    powershell -NoProfile $Current_Folder\Sources\Run_in_Sandbox\RunInSandbox_Config.ps1
}


Get-Config
Write-Progress -Activity $Progress_Activity -PercentComplete 10

Create-Checkpoint
Write-Progress -Activity $Progress_Activity -PercentComplete 20

Write-LogMessage -Message_Type "INFO" -Message "Adding context menu"
Write-LogMessage -Message_Type "INFO" -Message "OS version is: $Windows_Version"


if ($Add_CMD -eq $True) {
    Add-RegItem -Sub_Reg_Path "cmdfile" -Type "CMD"
    Add-RegItem -Sub_Reg_Path "batfile" -Type "BAT"
}
Write-Progress -Activity $Progress_Activity -PercentComplete 25

if ($Add_EXE -eq $True) {
    Add-RegItem -Sub_Reg_Path "exefile" -Type "EXE"
}
Write-Progress -Activity $Progress_Activity -PercentComplete 30

if ($Add_Folder -eq $True) {
    Add-RegItem -Sub_Reg_Path "Directory\Background" -Type "Folder_Inside" -Entry_Name "this folder" -Key_Label "Share this folder in a Sandbox"
    Add-RegItem -Sub_Reg_Path "Directory" -Type "Folder_On" -Entry_Name "this folder" -Key_Label "Share this folder in a Sandbox"
}
Write-Progress -Activity $Progress_Activity -PercentComplete 35

if ($Add_HTML -eq $True) {
    Add-RegItem -Sub_Reg_Path "MSEdgeHTM" -Type "HTML" -Key_Label "Run this web link in Sandbox"
    Add-RegItem -Sub_Reg_Path "ChromeHTML" -Type "HTML" -Key_Label "Run this web link in Sandbox"
    Add-RegItem -Sub_Reg_Path "IE.AssocFile.HTM" -Type "HTML" -Key_Label "Run this web link in Sandbox"
    Add-RegItem -Sub_Reg_Path "IE.AssocFile.URL" -Type "HTML" -Key_Label "Run this URL in Sandbox"
}
Write-Progress -Activity $Progress_Activity -PercentComplete 40

if ($Add_Intunewin -eq $True) {
    #Add-RegItem -Sub_Reg_Path ".intunewin" -Type "Intunewin"
    Add-RegItem -Sub_Reg_Path "SystemFileAssociations\.intunewin" -Type "Intunewin"
}
Write-Progress -Activity $Progress_Activity -PercentComplete 45

if ($Add_ISO -eq $True) {
    Add-RegItem -Sub_Reg_Path "Windows.IsoFile" -Type "ISO" -Key_Label "Extract ISO file in Sandbox"
    Add-RegItem -Reg_Path "$HKCU_Classes" -Sub_Reg_Path ".iso" -Type "ISO" -Key_Label "Extract ISO file in Sandbox"
}
Write-Progress -Activity $Progress_Activity -PercentComplete 50

if ($Add_MSI -eq $True) {
    Add-RegItem -Sub_Reg_Path "Msi.Package" -Type "MSI"
}
Write-Progress -Activity $Progress_Activity -PercentComplete 55

if ($Add_MSIX -eq $True) {
    $MSIX_Shell_Registry_Key = "Registry::HKEY_CLASSES_ROOT\.msix\OpenWithProgids"
    if (Test-Path -Path $MSIX_Shell_Registry_Key) {
        $Get_Default_Value = (Get-Item -Path $MSIX_Shell_Registry_Key).Property
        if ($Get_Default_Value) {
            Add-RegItem -Sub_Reg_Path "$Get_Default_Value" -Type "MSIX"
        } 
    }
    $Default_MSIX_HKCU = "$HKCU_Classes\.msix"
    if (Test-Path -Path $Default_MSIX_HKCU) {
        $Get_Default_Value = (Get-Item -Path "$Default_MSIX_HKCU\OpenWithProgids").Property
        if ($Get_Default_Value) {
            Add-RegItem -Reg_Path $HKCU_Classes -Sub_Reg_Path "$Get_Default_Value" -Type "MSIX"
        }
    }
}
Write-Progress -Activity $Progress_Activity -PercentComplete 60

if ($Add_MultipleApp -eq $True) {
    Add-RegItem -Sub_Reg_Path ".sdbapp" -Type "SDBApp" -Entry_Name "application bundle"
}
Write-Progress -Activity $Progress_Activity -PercentComplete 65

if ($Add_PDF -eq $True) {
    Add-RegItem -Sub_Reg_Path "SystemFileAssociations\.pdf" -Type "PDF" -Key_Label "Open PDF in Sandbox"
}
Write-Progress -Activity $Progress_Activity -PercentComplete 75

if ($Add_PPKG -eq $True) {
    Add-RegItem -Sub_Reg_Path "Microsoft.ProvTool.Provisioning.1" -Type "PPKG"
}
Write-Progress -Activity $Progress_Activity -PercentComplete 80

if ($Add_PS1 -eq $True) {
    Write-LogMessage -Message_Type "INFO" -Message "Checking OS Version for PS1..."
    if ($Windows_Version -like "*Windows 10*") {
        Write-LogMessage -Message_Type "INFO" -Message "Running on Windows 10"

        Add-RegItem -Sub_Reg_Path "SystemFileAssociations\.ps1" -Type "PS1Basic" -Entry_Name "PS1 as user" -Info_Type "PS1" -MainMenuLabel "Run PS1 in Sandbox" -MainMenuSwitch
        Add-RegItem -Sub_Reg_Path "SystemFileAssociations\.ps1" -Type "PS1System" -Entry_Name "PS1 as system" -Info_Type "PS1" -MainMenuLabel "Run PS1 in Sandbox" -MainMenuSwitch
        Add-RegItem -Sub_Reg_Path "SystemFileAssociations\.ps1" -Type "PS1Params" -Entry_Name "PS1 with Parameters" -Info_Type "PS1" -MainMenuLabel "Run PS1 in Sandbox" -MainMenuSwitch
    }
    
    if ($Windows_Version -like "*Windows 11*") {
        $Registry_Set = $False
        Write-LogMessage -Message_Type "INFO" -Message "Running on Windows 11"

        if (Test-Path -Path $HKCU_Classes) {
            $Default_PS1_HKCU = "$HKCU_Classes\.ps1"
            $OpenWithProgids_Key = "$Default_PS1_HKCU\OpenWithProgids"
            if (Test-Path -Path $OpenWithProgids_Key) {
                $Get_OpenWithProgids_Default_Value = (Get-Item -Path $OpenWithProgids_Key).Property
                ForEach ($Prop in $Get_OpenWithProgids_Default_Value) {
                    Add-RegItem -Reg_Path "$HKCU_Classes" -Sub_Reg_Path "$Prop" -Type "PS1Basic" -Entry_Name "PS1 as user" -Info_Type "PS1" -MainMenuLabel "Run PS1 in Sandbox" -MainMenuSwitch
                    Add-RegItem -Reg_Path "$HKCU_Classes" -Sub_Reg_Path "$Prop" -Type "PS1System" -Entry_Name "PS1 as system" -Info_Type "PS1" -MainMenuLabel "Run PS1 in Sandbox" -MainMenuSwitch
                    Add-RegItem -Reg_Path "$HKCU_Classes" -Sub_Reg_Path "$Prop" -Type "PS1Params" -Entry_Name "PS1 with Parameters" -Info_Type "PS1" -MainMenuLabel "Run PS1 in Sandbox" -MainMenuSwitch
                }
                $Registry_Set = $True
            }

            # ADDING CONTEXT MENU DEPENDING OF THE USERCHOICE
            # The userchoice for PS1 is located in: HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.ps1\UserChoice
            $PS1_UserChoice = "$HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.ps1\UserChoice"
            $Get_UserChoice = (Get-ItemProperty -Path $PS1_UserChoice).ProgID

            $HKCR_UserChoice_Key = "Registry::HKEY_CLASSES_ROOT\$Get_UserChoice"
            $PS1_Shell_Registry_Key = "$HKCR_UserChoice_Key\Shell"
            if (Test-Path -Path $PS1_Shell_Registry_Key) {
                Add-RegItem -Sub_Reg_Path "$Get_UserChoice" -Type "PS1Basic" -Entry_Name "PS1 as user" -Info_Type "PS1" -MainMenuLabel "Run PS1 in Sandbox" -MainMenuSwitch
                Add-RegItem -Sub_Reg_Path "$Get_UserChoice" -Type "PS1System" -Entry_Name "PS1 as system" -Info_Type "PS1" -MainMenuLabel "Run PS1 in Sandbox" -MainMenuSwitch
                Add-RegItem -Sub_Reg_Path "$Get_UserChoice" -Type "PS1Params" -Entry_Name "PS1 with Parameters" -Info_Type "PS1" -MainMenuLabel "Run PS1 in Sandbox" -MainMenuSwitch
                $Registry_Set = $True
            }
        }
        if ($Registry_Set -eq $False) {
            Write-LogMessage -Message_Type "WARNING" -Message "Couldn´t set the correct registry keys. You probably don´t have any programs selected as default for .ps1 extension!"
            Write-LogMessage -Message_Type "WARNING" -Message "Will try anyway using the method for Windows 10"
            Add-RegItem -Sub_Reg_Path "SystemFileAssociations\.ps1" -Type "PS1Basic" -Entry_Name "PS1 as user" -Info_Type "PS1" -MainMenuLabel "Run PS1 in Sandbox" -MainMenuSwitch
            Add-RegItem -Sub_Reg_Path "SystemFileAssociations\.ps1" -Type "PS1System" -Entry_Name "PS1 as system" -Info_Type "PS1" -MainMenuLabel "Run PS1 in Sandbox" -MainMenuSwitch
            Add-RegItem -Sub_Reg_Path "SystemFileAssociations\.ps1" -Type "PS1Params" -Entry_Name "PS1 with Parameters" -Info_Type "PS1" -MainMenuLabel "Run PS1 in Sandbox" -MainMenuSwitch
        }
    }
}
Write-Progress -Activity $Progress_Activity -PercentComplete 85

if ($Add_Reg -eq $True) {
    Add-RegItem -Sub_Reg_Path "regfile" -Type "REG" -Key_Label "Test reg file in Sandbox"
}
Write-Progress -Activity $Progress_Activity -PercentComplete 90

if ($Add_VBS -eq $True) {
    Add-RegItem -Sub_Reg_Path "VBSFile" -Type "VBSBasic" -Entry_Name "VBS" -MainMenuLabel "Run VBS in Sandbox" -MainMenuSwitch
    Add-RegItem -Sub_Reg_Path "VBSFile" -Type "VBSParams" -Entry_Name "VBS with Parameters" -Info_Type "VBS" -MainMenuLabel "Run VBS in Sandbox" -MainMenuSwitch
}
Write-Progress -Activity $Progress_Activity -PercentComplete 95

if ($Add_ZIP -eq $True) {
    # Run on ZIP
    Add-RegItem -Sub_Reg_Path "CompressedFolder" -Type "ZIP" -Key_Label "Extract ZIP in Sandbox"

    # Run on ZIP if WinRAR is installed
    if (Test-Path -Path "Registry::HKEY_CLASSES_ROOT\WinRAR.ZIP") {
        Add-RegItem -Sub_Reg_Path "WinRAR.ZIP" -Type "ZIP" -Key_Label "Extract ZIP (WinRAR) in Sandbox"
    }
    
    # Run on 7z
    if (Test-Path -Path "Registry::HKEY_CLASSES_ROOT\Applications\7zFM.exe") {
        Add-RegItem -Sub_Reg_Path "Applications\7zFM.exe" -Type "7z" -Info_Type "7z" -Entry_Name "ZIP" -Key_Label "Extract 7z file in Sandbox"
    }
    if (Test-Path -Path "Registry::HKEY_CLASSES_ROOT\7-Zip.7z") {
        Add-RegItem -Sub_Reg_Path "7-Zip.7z" -Type "7z" -Info_Type "7z" -Entry_Name "ZIP" -Key_Label "Extract 7z file in Sandbox"
    }
    if (Test-Path -Path "Registry::HKEY_CLASSES_ROOT\7-Zip.rar") {
        Add-RegItem -Sub_Reg_Path "7-Zip.rar" -Type "7z" -Info_Type "7z" -Entry_Name "ZIP" -Key_Label "Extract RAR file in Sandbox"
    }
}
Write-Progress -Activity $Progress_Activity -PercentComplete 100

Copy-Item -Path $Log_File -Destination $Destination_folder -Force