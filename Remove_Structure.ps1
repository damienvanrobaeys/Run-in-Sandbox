param (
    [Switch]$DeepClean
)

$Current_Folder = $PSScriptRoot

Unblock-File -Path $Current_Folder\CommonFunctions.ps1
. "$Current_Folder\CommonFunctions.ps1"


CheckFor-SandboxFolder
CheckFor-Admin
Get-Config

if ( (Test-Path -LiteralPath $Run_in_Sandbox_Folder) -and (-not $DeepClean) ) {
    Write-LogMessage -Message_Type "Warning" -Message "A lot of things have changed regarding installing and uninstalling Run-in-Sandbox"
    Write-LogMessage -Message_Type "Warning" -Message "It is recommended to run this script again with the -DeepClean parameter"
    Write-LogMessage -Message_Type "Warning" -Message "Otherwise it should be safe to just continue uninstalling, but leftovers might remain on the system"
    Write-LogMessage -Message_Type "Info" -Message "Press `"Enter`" button to continue"
    Read-Host
}

if ($DeepClean) {
    Write-LogMessage -Message_Type "INFO" -Message "Script has been started with deep cleaning enabled. This might take a moment"
    [String[]] $results = @()
    $results = Find-RegistryIconPaths -rootRegistryPath 'HKEY_CLASSES_ROOT'
    $results += Find-RegistryIconPaths -rootRegistryPath 'HKEY_CLASSES_ROOT\SystemFileAssociations'
    $results += Find-RegistryIconPaths -rootRegistryPath $HKCU_Classes
    $results = $results | Where-Object { $_ -notlike "REGISTRY::HKEY_CLASSES_ROOT\SystemFileAssociations\SystemFileAssociations*" }
    $results = $results | Select-Object -Unique | Sort-Object
    
    foreach ($reg_path in $results) {
        try {
            # Get all child items and sort by depth (deepest first)
            Get-ChildItem -Path $reg_path -Recurse | Sort-Object { $_.PSPath.Split('\').Count } -Descending | Select-Object -Property PSPath -ExpandProperty PSPath | Remove-Item -Force -Confirm:$false -ErrorAction Stop

            # Remove the main registry path if it still exists
            if (Test-Path -Path $reg_path) {
                Remove-Item -LiteralPath $reg_path -Force -Recurse -Confirm:$false -ErrorAction Stop
            }

            Write-LogMessage -Message_Type "SUCCESS" -Message "Path: `"$reg_path`" has been removed"
        } catch {
            Write-LogMessage -Message_Type "ERROR" -Message "Path: `"$reg_path`" couldn´t be removed"
        } 
    }
    Write-LogMessage -Message_Type "INFO" -Message "Deep cleaning finished"
}



if ($Add_CMD -eq $True) {
    Remove-RegItem -Sub_Reg_Path "cmdfile" -Type "CMD"
    Remove-RegItem -Sub_Reg_Path "batfile" -Type "BAT"
}

if ($Add_EXE -eq $True) {
    Remove-RegItem -Sub_Reg_Path "exefile" -Type "EXE"
}

if ($Add_Folder -eq $True) {
    Remove-RegItem -Sub_Reg_Path "Directory\Background" -Type "Folder_Inside" -Entry_Name "this folder" -Key_Label "Share this folder in a Sandbox"
    Remove-RegItem -Sub_Reg_Path "Directory" -Type "Folder_On" -Entry_Name "this folder" -Key_Label "Share this folder in a Sandbox"
}

if ($Add_HTML -eq $True) {
    Remove-RegItem -Sub_Reg_Path "MSEdgeHTM" -Type "HTML" -Key_Label "Run this web link in Sandbox"
    Remove-RegItem -Sub_Reg_Path "ChromeHTML" -Type "HTML" -Key_Label "Run this web link in Sandbox"
    Remove-RegItem -Sub_Reg_Path "IE.AssocFile.HTM" -Type "HTML" -Key_Label "Run this web link in Sandbox"
    Remove-RegItem -Sub_Reg_Path "IE.AssocFile.URL" -Type "HTML" -Key_Label "Run this URL in Sandbox"
}

if ($Add_Intunewin -eq $True) {
    Remove-RegItem -Sub_Reg_Path ".intunewin" -Type "Intunewin"
}

if ($Add_ISO -eq $True) {
    Remove-RegItem -Sub_Reg_Path "Windows.IsoFile" -Type "ISO" -Key_Label "Extract ISO file in Sandbox"
    Remove-RegItem -Reg_Path "$HKCU_Classes" -Sub_Reg_Path ".iso" -Type "ISO" -Key_Label "Extract ISO file in Sandbox"
}

if ($Add_MSI -eq $True) {
    Remove-RegItem -Sub_Reg_Path "Msi.Package" -Type "MSI"
}

if ($Add_MSIX -eq $True) {
    $MSIX_Shell_Registry_Key = "Registry::HKEY_CLASSES_ROOT\.msix\OpenWithProgids"
    if (Test-Path -Path $MSIX_Shell_Registry_Key) {
        $Get_Default_Value = (Get-Item -Path $MSIX_Shell_Registry_Key).Property
        if ($Get_Default_Value) {
            Remove-RegItem -Sub_Reg_Path "$Get_Default_Value" -Type "MSIX"
        } 
    }
    $Default_MSIX_HKCU = "$HKCU_Classes\.msix"
    if (Test-Path -Path $Default_MSIX_HKCU) {
        $Get_Default_Value = (Get-Item -Path "$Default_MSIX_HKCU\OpenWithProgids").Property
        if ($Get_Default_Value) {
            Remove-RegItem -Reg_Path $HKCU_Classes -Sub_Reg_Path "$Get_Default_Value" -Type "MSIX"
        }
    } 
}

if ($Add_MultipleApp -eq $True) {
    Remove-RegItem -Sub_Reg_Path ".sdbapp" -Type "SDBApp" -Entry_Name "application bundle"
}

if ($Add_PDF -eq $True) {
    Remove-RegItem -Sub_Reg_Path "SystemFileAssociations\.pdf" -Type "PDF" -Key_Label "Open PDF in Sandbox"
}

if ($Add_PPKG -eq $True) {
    Remove-RegItem -Sub_Reg_Path "Microsoft.ProvTool.Provisioning.1" -Type "PPKG"
}

if ($Add_PS1 -eq $True) {
    if ($Windows_Version -like "*Windows 10*") {
        Remove-RegItem -Sub_Reg_Path "SystemFileAssociations\.ps1" -Type "PS1"
    }
    
    if ($Windows_Version -like "*Windows 11*") {
        $Registry_Set = $False
        if (Test-Path $HKCU_Classes) {
            $Default_PS1_HKCU = "$HKCU_Classes\.ps1"
            
            $OpenWithProgids_Key = "$Default_PS1_HKCU\OpenWithProgids"
            if (Test-Path $OpenWithProgids_Key) {
                $Get_OpenWithProgids_Default_Value = (Get-Item $OpenWithProgids_Key).Property
                ForEach ($Prop in $Get_OpenWithProgids_Default_Value) {
                    Remove-RegItem -Reg_Path "$HKCU_Classes" -Sub_Reg_Path "$Prop" -Type "PS1"
                }
                $Registry_Set = $True
            }

            $PS1_UserChoice = "$HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.ps1\UserChoice"
            if (Test-Path -Path $PS1_UserChoice) {
                $Get_UserChoice = (Get-ItemProperty $PS1_UserChoice).ProgID
                $HKCR_UserChoice_Key = "Registry::HKEY_CLASSES_ROOT\$Get_UserChoice"
                Remove-RegItem -Sub_Reg_Path "$Get_UserChoice" -Type "PS1"
                $Registry_Set = $True
            }
        }
        if ($Registry_Set -eq $False) {
            Write-LogMessage -Message_Type "WARNING" -Message "Couldn´t remove the correct registry keys. You probably don´t have any programs selected as default for .ps1 extension!"
            Write-LogMessage -Message_Type "WARNING" -Message "Will try anyway using the method for Windows 10"
            Remove-RegItem -Sub_Reg_Path "SystemFileAssociations\.ps1" -Type "PS1"
        }
    }
}

if ($Add_Reg -eq $True) {
    Remove-RegItem -Sub_Reg_Path "regfile" -Type "REG" -Key_Label "Test reg file in Sandbox"
}

if ($Add_VBS -eq $True) {
    Remove-RegItem -Sub_Reg_Path "VBSFile" -Type "VBS"
}

if ($Add_ZIP -eq $True) {
    Remove-RegItem -Sub_Reg_Path "CompressedFolder" -Type "ZIP" -Key_Label "Extract ZIP in Sandbox"
    Remove-RegItem -Sub_Reg_Path "WinRAR.ZIP" -Type "ZIP" -Key_Label "Extract ZIP (WinRAR) in Sandbox"
    Remove-RegItem -Sub_Reg_Path "Applications\7zFM.exe" -Type "7z" -Info_Type "7z" -Entry_Name "ZIP" -Key_Label "Extract 7z file in Sandbox"
    Remove-RegItem -Sub_Reg_Path "7-Zip.7z" -Type "7z" -Info_Type "7z" -Entry_Name "ZIP" -Key_Label "Extract 7z file in Sandbox"
    Remove-RegItem -Sub_Reg_Path "7-Zip.rar" -Type "7z" -Info_Type "7z" -Entry_Name "ZIP" -Key_Label "Extract RAR file in Sandbox"
}

if (Test-Path -Path $Run_in_Sandbox_Folder) {
    try {
        Remove-Item $Run_in_Sandbox_Folder -Recurse -Force
        Write-LogMessage -Message_Type "Success" -Message "Run-in-Sandbox has been removed"
    } catch {
        Write-LogMessage -Message_Type "ERROR" -Message "Run-in-Sandbox Folder couldnt be removed"
        Write-LogMessage -Message_Type "INFO" -Message "Please remove path `"$Run_in_Sandbox_Folder`" manually"
    }
}
