#***************************************************************************************************************
# Author: Damien VAN ROBAEYS
# Website: http://www.systanddeploy.com
# Twitter: https://twitter.com/syst_and_deploy
# Purpose: This script will remove context menus added to run quickly files in Windows Sandbox
#***************************************************************************************************************

Function Write-LogMessage([string]$Message, [string]$Message_Type) {
    $MyDate = "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
    Write-Output "$MyDate - $Message_Type : $Message"
}

Function Remove-RegItem {
    param (
        $Reg_Path
    )

    Write-LogMessage -Message_Type "INFO" -Message "Removing registry path: $Reg_Path"

    if (Test-Path $Reg_Path) {
        Write-LogMessage -Message_Type "SUCCESS" -Message "Following registry path exists: $Reg_Path"
        try {
            Remove-Item -Path $Reg_Path -Recurse
            Write-LogMessage -Message_Type "SUCCESS" -Message "$Reg_Path has been removed"
        } catch {
            Write-LogMessage -Message_Type "ERROR" -Message "$Reg_Path has not been removed"
        }
    } else {
        Write-LogMessage -Message_Type "INFO" -Message "Can not find registry path: $Reg_Path"
    }
}


$ProgData = $env:ProgramData
$Sandbox_Folder = "$ProgData\Run_in_Sandbox"

$Sandbox_Folder = "$ProgData\Run_in_Sandbox"
if (-not (Test-Path $Sandbox_Folder) ) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [System.Windows.Forms.MessageBox]::Show("Can not find the folder $Sandbox_Folder")
    break
}
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
$Run_As_Admin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if ($Run_As_Admin -eq $False) {
    Write-LogMessage -Message_Type "ERROR" -Message "The script has not been lauched with admin rights"
    [System.Windows.Forms.MessageBox]::Show("Please run the tool with admin rights :-)")
    break
}
Write-LogMessage -Message_Type "INFO" -Message "The script has been launched with admin rights"

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
$Add_CMD = $Get_XML_Content.Configuration.ContextMenu_CMD
$Add_PDF = $Get_XML_Content.Configuration.ContextMenu_PDF

$List_Drive = Get-PSDrive | Where-Object { $_.Name -eq "HKCR_SD" }
if ($null -ne $List_Drive) { Remove-PSDrive $List_Drive }
New-PSDrive -PSProvider registry -Root HKEY_CLASSES_ROOT -Name HKCR_SD | Out-Null


if ($Add_PS1 -eq $True) {
    # REMOVE RUN ON PS1
    Write-Output "Removing context menu for PS1"

    $PS1_Main_Menu = "Run PS1 in Sandbox"
    $Windows_Version = (Get-CimInstance -class Win32_OperatingSystem).Caption
    if ($Windows_Version -like "*Windows 10*") {
        $PS1_Shell_Registry_Key = "HKCR_SD:\SystemFileAssociations\.ps1\Shell"

        if (Test-Path "$PS1_Shell_Registry_Key\$PS1_Main_Menu") {
            Remove-RegItem -Reg_Path "$PS1_Shell_Registry_Key\$PS1_Main_Menu"
        }
    }
    if ($Windows_Version -like "*Windows 11*") {
        $Current_User_SID = (Get-ChildItem Registry::\HKEY_USERS | Where-Object { Test-Path "$($_.pspath)\Volatile Environment" } | ForEach-Object { (Get-ItemProperty "$($_.pspath)\Volatile Environment") }).PSParentPath.split("\")[-1]																			# RUN ON ISO
        $HKCU_Classes = "Registry::HKEY_USERS\$Current_User_SID" + "_Classes"
        if (Test-Path $HKCU_Classes) {
            $Default_PS1_HKCU = "$HKCU_Classes\.ps1"
            $rOpenWithProgids_Key = "$Default_PS1_HKCU\rOpenWithProgids"
            if (Test-Path $rOpenWithProgids_Key) {
                $Get_rOpenWithProgids_Default_Value = (Get-Item "$Default_PS1_HKCU\rOpenWithProgids").Property
                ForEach ($Prop in $Get_rOpenWithProgids_Default_Value) {
                    $Default_HKCU_PS1_Shell_Registry_Key = "$HKCU_Classes\$Prop\Shell"
                    if (Test-Path $Default_HKCU_PS1_Shell_Registry_Key) {
                        $Main_Menu_Path = "$Default_HKCU_PS1_Shell_Registry_Key\$PS1_Main_Menu"
                        Remove-RegItem -Reg_Path "$Main_Menu_Path"
                    }
                }
            }

            $OpenWithProgids_Key = "$Default_PS1_HKCU\OpenWithProgids"
            if (Test-Path $OpenWithProgids_Key) {
                $Get_OpenWithProgids_Default_Value = (Get-Item $OpenWithProgids_Key).Property
                ForEach ($Prop in $Get_OpenWithProgids_Default_Value) {
                    $Default_HKCU_PS1_Shell_Registry_Key = "$HKCU_Classes\$Prop\Shell"
                    if (Test-Path $Default_HKCU_PS1_Shell_Registry_Key) {
                        $Main_Menu_Path = "$Default_HKCU_PS1_Shell_Registry_Key\$PS1_Main_Menu"
                        Remove-RegItem -Reg_Path "$Main_Menu_Path"
                    }
                }
            }

            # RE%OVING CONTEXT MENU DEPENDING OF THE USERCHOICE
            $Current_User_SID = (Get-ChildItem Registry::\HKEY_USERS | Where-Object { Test-Path "$($_.pspath)\Volatile Environment" } | ForEach-Object { (Get-ItemProperty "$($_.pspath)\Volatile Environment") }).PSParentPath.split("\")[-1]																			# RUN ON ISO
            $HKCU = "Registry::HKEY_USERS\$Current_User_SID"
            $HKCU_Classes = "Registry::HKEY_USERS\$Current_User_SID" + "_Classes"
            if (Test-Path $HKCU) {
                $PS1_UserChoice = "$HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.ps1\UserChoice"
                $Get_UserChoice = (Get-ItemProperty $PS1_UserChoice).ProgID
                $HKCR_UserChoice_Key = "HKCR_SD:\$Get_UserChoice"
                $HKCR_UserChoice_Shell = "$HKCR_UserChoice_Key\Shell"
                if (Test-Path $HKCR_UserChoice_Shell) {
                    $HKCR_UserChoice_Label = "$HKCR_UserChoice_Shell\$PS1_Main_Menu"
                    if (Test-Path $HKCR_UserChoice_Label) {
                        Remove-RegItem -Reg_Path $HKCR_UserChoice_Label
                    }
                }
            }
        }
    }
}

if ($Add_Reg -eq $True) {
    # REMOVE RUN ON REG
    Write-Output "Removing context menu for REG"
    $Reg_Shell_Registry_Key = "HKCR_SD:\regfile\Shell"
    $Reg_Key_Label = "Test reg file in Sandbox"
    Remove-RegItem -Reg_Path "$REG_Shell_Registry_Key\$Reg_Key_Label"

    if (Test-Path "$REG_Shell_Registry_Key\Test the reg file in Sandbox") {
        Remove-RegItem -Reg_Path "$REG_Shell_Registry_Key\Test the reg file in Sandbox"
    }
}

if ($Add_ISO -eq $True) {
    $ISO_Key_Label = "Extract ISO file in Sandbox"

    # REMOVE RUN ON REG from HKCR under Windows.IsoFile
    Write-Output "Removing context menu for ISO"
    $ISO_Shell_Registry_Key = "HKCR_SD:\Windows.IsoFile\Shell"
    if (Test-Path "$ISO_Shell_Registry_Key\$ISO_Key_Label") {
        Remove-RegItem -Reg_Path "$ISO_Shell_Registry_Key\$ISO_Key_Label"
    }

    $ISO_Key = "HKCR_SD:\.ISO"
    if (Test-Path $ISO_Key) {
        Write-LogMessage -Message_Type "INFO" -Message "The key HKCR\.ISO exists"
        $Get_ISO_Keys = Get-Item $ISO_Key
        ForEach ($Key in $Get_ISO_Keys) {
            $Get_Properties = $Key.Property
            Write-LogMessage -Message_Type "INFO" -Message "Following subkeys found: $Get_Properties"
            foreach ($Property in $Get_Properties) {
                $Prop = (Get-ItemProperty $ISO_Key)."$Property"
                Write-LogMessage -Message_Type "INFO" -Message "Following property found: $Prop"
                $ISO_Property_Key = "$HKCR_SD\$Prop"
                Write-LogMessage -Message_Type "INFO" -Message "Reg path to test: $ISO_Property_Key"
                if (Test-Path $ISO_Property_Key) {
                    Write-LogMessage -Message_Type "INFO" -Message "The following reg path exists: $ISO_Property_Key"
                    $ISO_Property_Shell = "$ISO_Property_Key\Shell"
                    if (Test-Path $ISO_Property_Shell) {
                        Write-LogMessage -Message_Type "INFO" -Message "The following reg path exists: $ISO_Property_Shell"
                        $ISO_Key_Label_Path = "$ISO_Property_Shell\$ISO_Key_Label"
                        if (Test-Path $ISO_Key_Label_Path) {
                            Remove-RegItem -Reg_Path $ISO_Key_Label_Path
                        }
                    }
                } else {
                    Write-LogMessage -Message_Type "INFO" -Message "The following reg path does not exist: $ISO_Property_Key"
                }
            }
        }
    }



    # REMOVE RUN ON REG from HKCU if 7zip exists
    $ISO_Shell_HKCU_Registry_Key = "Registry::HKEY_USERS\$Current_User_SID"
    $HKCU_Classes = "$ISO_Shell_HKCU_Registry_Key\SOFTWARE\Classes"
    $Default_ISO_HKCU = "$HKCU_Classes\.iso"
    if (Test-Path $Default_ISO_HKCU) {
        $Get_Default_Value = (Get-ItemProperty $Default_ISO_HKCU)."(default)"
        if ($Get_Default_Value -eq "7-Zip.iso") {
            $Default_HKCU_ISO_Shell_Registry_Key = "$HKCU_Classes\$Get_Default_Value\Shell"
            $ISO_HKCU_Key_Label_Path = "$Default_HKCU_ISO_Shell_Registry_Key\$ISO_Key_Label"
            if (Test-Path $ISO_HKCU_Key_Label_Path) {
                Remove-RegItem -Reg_Path $ISO_HKCU_Key_Label_Path
            }
        }
    }
}

if ($Add_MSIX -eq $True) {
    Write-Output "Removing context menu for MSIX"
    $MSIX_Key_Label = "Run MSIX file in Sandbox"
    # REMOVE RUN ON REG from HKCR
    $MSIX_Shell_Registry_Key = "HKCR_SD:\.msix\OpenWithProgids"
    if (Test-Path $MSIX_Shell_Registry_Key) {
        $Get_Default_Value = (Get-Item $MSIX_Shell_Registry_Key).Property
        $MSIX_Shell_Registry = "HKCR_SD:\$Get_Default_Value\Shell"
        if (Test-Path $MSIX_Shell_Registry) {
            $MSIX_Key_Label_Path = "$MSIX_Shell_Registry\$MSIX_Key_Label"
            if (Test-Path $MSIX_Key_Label_Path) {
                Remove-RegItem -Reg_Path $MSIX_Key_Label_Path
                Remove-RegItem -Reg_Path "$MSIX_Shell_Registry\$MSIX_Key_Label"
            }
        }
    }

    # Modify value from HKCU
    $Current_User_SID = (Get-ChildItem Registry::\HKEY_USERS | Where-Object { Test-Path "$($_.pspath)\Volatile Environment" } | ForEach-Object { (Get-ItemProperty "$($_.pspath)\Volatile Environment") }).PSParentPath.split("\")[-1]																			# RUN ON ISO
    $HKCU_Classes = "Registry::HKEY_USERS\$Current_User_SID" + "_Classes"
    if (Test-Path $HKCU_Classes) {
        $Default_MSIX_HKCU = "$HKCU_Classes\.msix"
        $Get_Default_Value = (Get-Item "$Default_MSIX_HKCU\OpenWithProgids").Property
        $Default_HKCU_MSIX_Shell_Registry_Key = "$HKCU_Classes\$Get_Default_Value\Shell"
        $MSIX_HKCU_Key_Label_Path = "$Default_HKCU_MSIX_Shell_Registry_Key\$MSIX_Key_Label"
        if (Test-Path $MSIX_HKCU_Key_Label_Path) {
            Remove-RegItem -Reg_Path "$MSIX_HKCU_Key_Label_Path"
        }
    }
}

if ($Add_PPKG -eq $True) {
    # REMOVE RUN ON PPKG
    Write-Output "Removing context menu for PPKG"
    $PPKG_Shell_Registry_Key = "HKCR_SD:\Microsoft.ProvTool.Provisioning.1\Shell"
    $PPKG_Key_Label = "Run PPKG file in Sandbox"
    Remove-RegItem -Reg_Path "$PPKG_Shell_Registry_Key\$PPKG_Key_Label"
}

if ($Add_HTML -eq $True) {
    $HTML_Key_Label = "Run this web link in Sandbox"

    # RUN ON HTML for Edge
    $HTML_Edge_Shell_Registry_Key = "HKCR_SD:\MSEdgeHTM\Shell"
    Remove-RegItem -Reg_Path "$HTML_Edge_Shell_Registry_Key\$HTML_Key_Label"

    # RUN ON HTML for Chrome
    $HTML_Chrome_Shell_Registry_Key = "HKCR_SD:\ChromeHTML\Shell"
    Remove-RegItem -Reg_Path "$HTML_Chrome_Shell_Registry_Key\$HTML_Key_Label"

    # RUN ON HTML for IE
    $HTML_IE_Shell_Registry_Key = "HKCR_SD:\IE.AssocFile.HTM\Shell"
    Remove-RegItem -Reg_Path "$HTML_IE_Shell_Registry_Key\$HTML_Key_Label"

    # RUN ON URL
    $URL_Shell_Registry_Key = "HKCR_SD:\IE.AssocFile.URL\Shell"
    $URL_Key_Label_Path = "Run this URL in Sandbox"
    # $URL_Key_Label_Path = "Run this web link in Sandbox"
    Remove-RegItem -Reg_Path "$URL_Shell_Registry_Key\$URL_Key_Label_Path"
}

if ($Add_EXE -eq $True) {
    # REMOVE RUN ON EXE
    Write-Output "Removing context menu for PS1"
    $EXE_Shell_Registry_Key = "HKCR_SD:\exefile\Shell"
    $EXE_Basic_Run = "Run EXE in Sandbox"
    Remove-RegItem -Reg_Path "$EXE_Shell_Registry_Key\$EXE_Basic_Run"

    if (Test-Path "$EXE_Shell_Registry_Key\Run the EXE in Sandbox") {
        Remove-RegItem -Reg_Path "$EXE_Shell_Registry_Key\Run the EXE in Sandbox"
    }
}

if ($Add_MSI -eq $True) {
    # RUN ON MSI
    Write-Output "Removing context menu for MSI"
    $MSI_Shell_Registry_Key = "HKCR_SD:\Msi.Package\Shell"
    $MSI_Basic_Run = "Run MSI in Sandbox"
    Remove-RegItem -Reg_Path "$MSI_Shell_Registry_Key\$MSI_Basic_Run"

    if (Test-Path "$MSI_Shell_Registry_Key\Run the MSI in Sandbox") {
        Remove-RegItem -Reg_Path "$MSI_Shell_Registry_Key\Run the MSI in Sandbox"
    }
}

if ($Add_Folder -eq $True) {
    Write-Output "Removing context menu for folder"
    # Share this folder - Inside the folder
    $Folder_Inside_Shell_Registry_Key = "HKCR_SD:\Directory\Background\shell"
    $Folder_Inside_Basic_Run = "Share this folder in a Sandbox"
    Remove-RegItem -Reg_Path "$Folder_Inside_Shell_Registry_Key\$Folder_Inside_Basic_Run"

    # Share this folder - Right-click on the folder
    $Folder_On_Shell_Registry_Key = "HKCR_SD:\Directory\shell"
    $Folder_On_Run = "Share this folder in a Sandbox"
    Remove-RegItem -Reg_Path "$Folder_On_Shell_Registry_Key\$Folder_On_Run"
}

if ($Add_Intunewin -eq $True) {
    # RUN ON Intunewin
    Write-Output "Removing context menu for intunewin"
    Remove-RegItem -Reg_Path "HKCR_SD:\.intunewin"
}

if ($Add_MultipleApp -eq $True) {
    # RUN ON multiple app context menu
    Write-Output "Removing context menu for multiple app"
    Remove-RegItem -Reg_Path "HKCR_SD:\.sdbapp"
}

if ($Add_VBS -eq $True) {
    # REMOVE RUN ON VBS
    Write-Output "Removing context menu for VBS"
    $VBS_Shell_Registry_Key = "HKCR_SD:\VBSFile\Shell"
    $VBS_Basic_Run = "Run VBS in Sandbox"
    $VBS_Parameter_Run = "Run VBS in Sandbox with parameters"
    Remove-RegItem -Reg_Path "$VBS_Shell_Registry_Key\$VBS_Basic_Run"
    Remove-RegItem -Reg_Path "$VBS_Shell_Registry_Key\$VBS_Parameter_Run"

    if (Test-Path "$VBS_Shell_Registry_Key\Run the VBS in Sandbox") {
        Remove-RegItem -Reg_Path "$VBS_Shell_Registry_Key\Run the VBS in Sandbox"
    }

    if (Test-Path "$VBS_Shell_Registry_Key\Run the VBS in Sandbox") {
        Remove-RegItem -Reg_Path "$VBS_Shell_Registry_Key\Run the VBS in Sandbox with parameters"
    }
}

if ($Add_ZIP -eq $True) {
    Write-Output "Removing context menu for ZIP"
    # RUN ON ZIP
    $ZIP_Shell_Registry_Key = "HKCR_SD:\CompressedFolder\Shell"
    $ZIP_Basic_Run = "Extract ZIP in Sandbox"
    Remove-RegItem -Reg_Path "$ZIP_Shell_Registry_Key\$ZIP_Basic_Run"

    if (Test-Path "$ZIP_Shell_Registry_Key\Extract ZIP in Sandbox") {
        Remove-RegItem -Reg_Path "$ZIP_Shell_Registry_Key\Extract ZIP in Sandbox"
    }

    # RUN ON ZIP if WinRAR is installed
    if (Test-Path "HKCR_SD:\WinRAR.ZIP\Shell\Extract RAR file in Sandbox") {
        $ZIP_WinRAR_Shell_Registry_Key = "HKCR_SD:\WinRAR.ZIP\Shell"
        # Remove-RegItem -Reg_Path "$ZIP_WinRAR_Shell_Registry_Key\$ZIP_Basic_Run"
        Remove-RegItem -Reg_Path "$ZIP_WinRAR_Shell_Registry_Key\Extract ZIP (WinRAR) in Sandbox"
    }


    # RAR with 7z
    if (Test-Path "HKCR_SD:\WinRAR.ZIP\Shell\Extract 7z file in Sandbox") {
        $Shell_Registry_Key = "HKCR_SD:\Applications\7zFM.exe\Shell"
        # Remove-RegItem -Reg_Path "$ZIP_WinRAR_Shell_Registry_Key\$ZIP_Basic_Run"
        Remove-RegItem -Reg_Path "$Shell_Registry_Key\Extract 7z file in Sandbox"
    }

    if (Test-Path "HKCR_SD:\WinRAR.ZIP\Shell\Extract 7z file in Sandbox") {
        $Shell_Registry_Key = "HKCR_SD:\7-Zip.7z\Shell"
        # Remove-RegItem -Reg_Path "$ZIP_WinRAR_Shell_Registry_Key\$ZIP_Basic_Run"
        Remove-RegItem -Reg_Path "$Shell_Registry_Key\Extract 7z file in Sandbox"
    }

    # RAR with 7z
    if (Test-Path "HKCR_SD:\WinRAR.ZIP\Shell\Extract RAR file in Sandbox") {
        $Shell_Registry_Key = "HKCR_SD:\7-Zip.rar\Shell"
        # Remove-RegItem -Reg_Path "$ZIP_WinRAR_Shell_Registry_Key\$ZIP_Basic_Run"
        Remove-RegItem -Reg_Path "$Shell_Registry_Key\Extract RAR file in Sandbox"
    }

    # REMOVE RUN ON 7Z
    $7z_Key_Label = "Extract 7z file in Sandbox"
    $7z_Shell_Registry_Key = "HKCR_SD:\.7z"
    if (Test-Path $7z_Shell_Registry_Key) {
        $Get_Default_Value = (Get-ItemProperty "HKCR_SD:\.7z")."(default)"
        $Default_ZIP_Shell_Registry_Key = "HKCR_SD:\$Get_Default_Value\Shell"
        if (Test-Path $Default_ZIP_Shell_Registry_Key) {
            if (Test-Path $Default_ZIP_Shell_Registry_Key) {
                Remove-RegItem -Reg_Path "$Default_ZIP_Shell_Registry_Key\$7z_Key_Label"
            }
        }
    }
    
    $7z_Key_Label = "Extract 7z file in Sandbox"
    $7z_Shell_Registry_Key = "HKCR_SD:\Applications\7zFM.exe"
    if (Test-Path $7z_Shell_Registry_Key) {
        $Default_ZIP_Shell_Registry_Key = "HKCR_SD:\Applications\7zFM.exe\Shell"
        if (Test-Path $Default_ZIP_Shell_Registry_Key) {
            Remove-RegItem -Reg_Path "$Default_ZIP_Shell_Registry_Key\$7z_Key_Label"
        }
    }

    # Checking default zip from HKCU
    $Current_User_SID = (Get-ChildItem Registry::\HKEY_USERS | Where-Object { Test-Path "$($_.pspath)\Volatile Environment" } | ForEach-Object { (Get-ItemProperty "$($_.pspath)\Volatile Environment") }).PSParentPath.split("\")[-1]																			# RUN ON ISO
    $HKCU = "Registry::HKEY_USERS\$Current_User_SID"
    $HKCU_Classes = "Registry::HKEY_USERS\$Current_User_SID" + "_Classes"
    if (Test-Path $HKCU) {
        $ZIP_UserChoice = "$HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\OpenWithProgids"
        $Get_Properties = (Get-Item $ZIP_UserChoice).property
        ForEach ($Prop in $Get_Properties) {
            $HKCR_UserChoice_Key = "HKCR_SD:\$Prop"
            $HKCR_UserChoice_Key
            $HKCR_UserChoice_Shell = "$HKCR_UserChoice_Key\Shell"
            if (Test-Path $HKCR_UserChoice_Shell) {
                $HKCR_UserChoice_Label = "$HKCR_UserChoice_Shell\$ZIP_Basic_Run"
                if (!(Test-Path $HKCR_UserChoice_Label)) {
                    Remove-RegItem -Reg_Path $HKCR_UserChoice_Label
                }
            }
        }
    }
}


if ($Add_CMD -eq $True) {
    # REMOVE RUN ON CMD
    Write-Output "Removing context menu for CMD"
    $CMD_Shell_Registry_Key = "HKCR_SD:\cmdfile\Shell"
    $CMD_Key_Label = "Run CMD in Sandbox"

    if (Test-Path "$CMD_Shell_Registry_Key\$CMD_Key_Label") {
        Remove-RegItem -Reg_Path "$CMD_Shell_Registry_Key\$CMD_Key_Label"
    }

    # REMOVE RUN ON BAT
    Write-Output "Removing context menu for BAT"
    $BAT_Shell_Registry_Key = "HKCR_SD:\batfile\Shell"
    $BAT_Key_Label = "Run BAT in Sandbox"

    if (Test-Path "$BAT_Shell_Registry_Key\$BAT_Key_Label") {
        Remove-RegItem -Reg_Path "$BAT_Shell_Registry_Key\$BAT_Key_Label"
    }
}

if ($Add_PDF -eq $True) {
    # REMOVE RUN ON CMD
    Write-Output "Removing context menu for PDF"
    $PDF_Shell_Registry_Key = "HKCR_SD:\SystemFileAssociations\.pdf\Shell"
    $PDF_Key_Label = "Open PDF in Sandbox"

    if (Test-Path "$PDF_Shell_Registry_Key\$PDF_Key_Label") {
        Remove-RegItem -Reg_Path "$PDF_Shell_Registry_Key\$PDF_Key_Label"
    }
}

if ($null -ne $List_Drive) { Remove-PSDrive $List_Drive }
Remove-Item $Sandbox_Folder -Recurse -Force


