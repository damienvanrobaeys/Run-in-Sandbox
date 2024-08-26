# Define global variables
$TEMP_Folder = $env:temp
$Log_File = "$TEMP_Folder\RunInSandbox_Install.log"
$Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"
$XML_Config = "$Run_in_Sandbox_Folder\Sandbox_Config.xml"
$Windows_Version = (Get-CimInstance -class Win32_OperatingSystem).Caption
$Current_User_SID = (Get-ChildItem -Path Registry::\HKEY_USERS | Where-Object { Test-Path -Path "$($_.pspath)\Volatile Environment" } | ForEach-Object { (Get-ItemProperty -Path "$($_.pspath)\Volatile Environment") }).PSParentPath.split("\")[-1]
$HKCU = "Registry::HKEY_USERS\$Current_User_SID"
$HKCU_Classes = "Registry::HKEY_USERS\$Current_User_SID" + "_Classes"
$Sandbox_Icon = "$env:ProgramData\Run_in_Sandbox\sandbox.ico"
$Sources = $Current_Folder + "\" + "Sources\*"
$Exported_Keys = @()

[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

# Function to write log messages
function Write-LogMessage {
    param (
        [string]$Message,
        [string]$Message_Type
    )

    $MyDate = "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
    Add-Content -Path $Log_File -Value "$MyDate - $Message_Type : $Message"
    $ForegroundColor = switch ($Message_Type) {
        "INFO"    { 'White' }
        "SUCCESS" { 'Green' }
        "WARNING" { 'Yellow' }
        "ERROR"   { 'DarkRed' }
        default   { 'White' }
    }
    Write-Host "$MyDate - $Message_Type : $Message" -ForegroundColor $ForegroundColor
}

# Function to export registry configuration
function Export-RegConfig {
    param (
        [string] $Reg_Path,
        [string] $Backup_Folder = "$Run_in_Sandbox_Folder\Registry_Backup",
        [string] $Type,
        [string] $Sub_Reg_Path
    )
    
    if ($Exported_Keys -contains $Reg_Path) {
        $Exported_Keys.Add($Reg_Path)
    } else {
        return
    }
    
    if (-not (Test-Path $Backup_Folder) ) {
        New-Item -ItemType Directory -Path $Backup_Folder -Force | Out-Null
    }
    
    Write-LogMessage -Message_Type "INFO" -Message "Exporting registry keys"
    
    $Backup_Path = $Backup_Folder + "\" + "Backup_" + $Type
    if ($Sub_Reg_Path) {
        $Backup_Path = $Backup_Path + "_" + $Sub_Reg_Path
    }
    $Backup_Path = $Backup_Path + ".reg"
    
    reg export $Reg_Path $Backup_Path /y > $null 2>&1

    # Check if the command ran successfully
    if ($?) {
        Write-LogMessage -Message_Type "SUCCESS" -Message "Exported `"$Reg_Path`" to `"$Backup_Path`""
    } else {
        Write-LogMessage -Message_Type "ERROR" -Message "Failed to export `"$Reg_Path`""
    }
}

# Function to add a registry item
function Add-RegItem {
    param (
        [string] $Reg_Path = "Registry::HKEY_CLASSES_ROOT",
        [string] $Sub_Reg_Path,
        [string] $Type,
        [string] $Entry_Name = $Type,
        [string] $Info_Type = $Type,
        [string] $Key_Label = "Run $Entry_Name in Sandbox",
        [string] $RegistryPathsFile = "$Run_in_Sandbox_Folder\RegistryEntries.txt",
        [string] $MainMenuLabel,
        [switch] $MainMenuSwitch
    )
    
    $Base_Registry_Key = "$Reg_Path\$Sub_Reg_Path"
    $Shell_Registry_Key = "$Base_Registry_Key\Shell"
    $Key_Label_Path = "$Shell_Registry_Key\$Key_Label"
    $MainMenuLabel_Path = "$Shell_Registry_Key\$MainMenuLabel"
    $Command_Path = "$Key_Label_Path\Command"
    $Command_for = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -WindowStyle Hidden -NoProfile -ExecutionPolicy Unrestricted -sta -File C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -Type $Type -ScriptPath `"%V`""
    
    Export-RegConfig -Reg_Path $($Base_Registry_Key.Split("::")[-1]) -Type $Type -Sub_Reg_Path $Sub_Reg_Path -ErrorAction Continue
    
    try {
        # Log the root registry path to the specified file
        if (-not (Test-Path $RegistryPathsFile) ) {
            New-Item -ItemType File -Path $RegistryPathsFile -Force | Out-Null
        }

        if (-not (Test-Path -Path $Base_Registry_Key) ) {
            New-Item -Path $Base_Registry_Key -ErrorAction Stop | Out-Null
        }
        
        if (-not (Test-Path -Path $Shell_Registry_Key) ) {
            New-Item -Path $Shell_Registry_Key -ErrorAction Stop | Out-Null
        }
        
        if ($MainMenuSwitch) {
            if ( -not (Test-Path $MainMenuLabel_Path) ) {
                New-Item -Path $Shell_Registry_Key -Name $MainMenuLabel -Force | Out-Null
                New-ItemProperty -Path $MainMenuLabel_Path -Name "subcommands" -PropertyType String | Out-Null
                New-Item -Path $MainMenuLabel_Path -Name "Shell" -Force | Out-Null
                New-ItemProperty -Path $MainMenuLabel_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon -ErrorAction Stop | Out-Null
            }
            $Key_Label_Path = "$MainMenuLabel_Path\Shell\$Key_Label"
            $Command_Path = "$Key_Label_Path\Command"
        }

        if (Test-Path -Path $Key_Label_Path) {
            Write-LogMessage -Message_Type "SUCCESS" -Message "Context menu for $Type has already been added"
            Add-Content -Path $RegistryPathsFile -Value $Key_Label_Path
            return
        }

        New-Item -Path $Key_Label_Path -ErrorAction Stop | Out-Null
        New-Item -Path $Command_Path -ErrorAction Stop | Out-Null
        if (-not $MainMenuSwitch) {
            New-ItemProperty -Path $Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon -ErrorAction Stop | Out-Null
        }
        Set-Item -Path $Command_Path -Value $Command_for -Force -ErrorAction Stop | Out-Null

        Add-Content -Path $RegistryPathsFile -Value $Key_Label_Path

        Write-LogMessage -Message_Type "SUCCESS" -Message "Context menu for `"$Info_Type`" has been added"
    } catch {
        Write-LogMessage -Message_Type "ERROR" -Message "Context menu for $Type could not be added"
    }
}

# Function to remove a registry item
function Remove-RegItem {
    param (
        [string] $Reg_Path = "Registry::HKEY_CLASSES_ROOT",
        [Parameter(Mandatory=$true)] [string] $Sub_Reg_Path,
        [Parameter(Mandatory=$true)] [string] $Type,
        [string] $Entry_Name = $Type,
        [string] $Info_Type = $Type,
        [string] $Key_Label = "Run $Entry_Name in Sandbox",
        [string] $MainMenuLabel,
        [switch] $MainMenuSwitch
    )
    Write-LogMessage -Message_Type "INFO" -Message "Removing context menu for $Type"
    $Base_Registry_Key = "$Reg_Path\$Sub_Reg_Path"
    $Shell_Registry_Key = "$Base_Registry_Key\Shell"
    $Key_Label_Path = "$Shell_Registry_Key\$Key_Label"
    
    
    if (-not (Test-Path -Path $Key_Label_Path) ) {
        if ($DeepClean) {
            Write-LogMessage -Message_Type "INFO" -Message "Registry Path for $Type has already been removed by deepclean"
            return
        }
        Write-LogMessage -Message_Type "WARNING" -Message "Could not find path for $Type"
        return
    }
    
    try {
        # Get all child items and sort by depth (deepest first)
        $ChildItems = Get-ChildItem -Path $Key_Label_Path -Recurse | Sort-Object { $_.PSPath.Split('\').Count } -Descending

        foreach ($ChildItem in $ChildItems) {
            Remove-Item -LiteralPath $ChildItem.PSPath -Force -ErrorAction Stop
        }

        # Remove the main registry path if it still exists
        if (Test-Path -Path $Key_Label_Path) {
            Remove-Item -LiteralPath $Key_Label_Path -Force -ErrorAction Stop
        }
        
        Write-LogMessage -Message_Type "SUCCESS" -Message "Context menu for `"$Info_Type`" has been removed"
    } catch {
        Write-LogMessage -Message_Type "ERROR" -Message "Context menu for $Type couldn´t be removed"
    }
}

function Find-RegistryIconPaths {
    param (
        [Parameter(Mandatory=$true)] [string]$rootRegistryPath,
        [string]$iconValueToMatch = "C:\\ProgramData\\Run_in_Sandbox\\sandbox.ico"
    )

    # Export the registry at the specified rootRegistryPath
    $exportPath = "$env:TEMP\registry_export.reg"
    reg export $rootRegistryPath $exportPath /y > $null 2>&1

    # Initialize an empty array to store matching paths
    $matchingPaths = @()

    # Read the exported registry file
    $lines = Get-Content -Path $exportPath

    # Process each line in the exported registry file
    foreach ($line in $lines) {
        # Check if the line defines a new key
        if ($line -match '^\[([^\]]+)\]$') {
            $currentPath = $matches[1]
        }

        # If the line contains the icon value, add the current path to the list
        # If the line contains the icon value, add the current path to the list
        if ($line -match '^\s*\"Icon\"=\"([^\"]+)\"$' -and $matches[1] -eq $iconValueToMatch) {
            $currentPath = "REGISTRY::$currentPath"
            $matchingPaths += $currentPath
        }
    }
    $matchingPaths = $matchingPaths | Sort-Object
    return $matchingPaths
}

# Function to get the configuration from XML
function Get-Config {
    if ( [string]::IsNullOrEmpty($XML_Config) ) {
        return
    }
    if (-not (Test-Path -Path $XML_Config) ) {
        return
    }
    $Get_XML_Content = [xml](Get-Content $XML_Config)
    
    $script:Add_EXE = $Get_XML_Content.Configuration.ContextMenu_EXE
    $script:Add_MSI = $Get_XML_Content.Configuration.ContextMenu_MSI
    $script:Add_PS1 = $Get_XML_Content.Configuration.ContextMenu_PS1
    $script:Add_VBS = $Get_XML_Content.Configuration.ContextMenu_VBS
    $script:Add_ZIP = $Get_XML_Content.Configuration.ContextMenu_ZIP
    $script:Add_Folder = $Get_XML_Content.Configuration.ContextMenu_Folder
    $script:Add_Intunewin = $Get_XML_Content.Configuration.ContextMenu_Intunewin
    $script:Add_MultipleApp = $Get_XML_Content.Configuration.ContextMenu_MultipleApp
    $script:Add_Reg = $Get_XML_Content.Configuration.ContextMenu_Reg
    $script:Add_ISO = $Get_XML_Content.Configuration.ContextMenu_ISO
    $script:Add_PPKG = $Get_XML_Content.Configuration.ContextMenu_PPKG
    $script:Add_HTML = $Get_XML_Content.Configuration.ContextMenu_HTML
    $script:Add_MSIX = $Get_XML_Content.Configuration.ContextMenu_MSIX
    $script:Add_CMD = $Get_XML_Content.Configuration.ContextMenu_CMD
    $script:Add_PDF = $Get_XML_Content.Configuration.ContextMenu_PDF
}

# Function to check if the script is run with admin privileges
function CheckFor-Admin {
    $Run_As_Admin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $Run_As_Admin) {
        Write-LogMessage -Message_Type "ERROR" -Message "The script has not been launched with admin rights"
        [System.Windows.Forms.MessageBox]::Show("Please run the tool with admin rights :-)")
        EXIT
    }
    Write-LogMessage -Message_Type "INFO" -Message "The script has been launched with admin rights"
}

# Function to check for source files
function CheckFor-Sources {
    if (-not (Test-Path -Path $Sources)) {
        Write-LogMessage -Message_Type "ERROR" -Message "Sources folder is missing"
        [System.Windows.Forms.MessageBox]::Show("It seems you haven´t downloaded all the folder structure.`nThe folder `"Sources`" is missing !!!")
        EXIT
    }
    Write-LogMessage -Message_Type "SUCCESS" -Message "The sources folder exists"
    
    $Check_Sources_Files_Count = (Get-ChildItem -Path "$Current_Folder\Sources\Run_in_Sandbox" -Recurse).count
    if ($Check_Sources_Files_Count -lt 40) {
        Write-LogMessage -Message_Type "ERROR" -Message "Some contents are missing"
        [System.Windows.Forms.MessageBox]::Show("It seems you haven´t downloaded all the folder structure !!!")
        EXIT
    }
}

# Function to check if the Windows Sandbox feature is installed
function CheckFor-Sandbox {
    try {
        $Is_Sandbox_Installed = (Get-WindowsOptionalFeature -Online -ErrorAction SilentlyContinue | Where-Object { $_.featurename -eq "Containers-DisposableClientVM" }).state
    } catch {
        if (Test-Path -Path "C:\Windows\System32\WindowsSandbox.exe") {
            Write-LogMessage -Message_Type "WARNING" -Message "It looks like you have the `Windows Sandbox` Feature installed, but your `TrustedInstaller` Service is disabled."
            Write-LogMessage -Message_Type "WARNING" -Message "The Script will continue, but you should check for issues running Windows Sandbox."
            $Is_Sandbox_Installed = "Enabled"
        } else {
            $Is_Sandbox_Installed = "Disabled"
        }
    }
    if ($Is_Sandbox_Installed -eq "Disabled") {
        Write-LogMessage -Message_Type "ERROR" -Message "The feature `Windows Sandbox` is not installed !!!"
        [System.Windows.Forms.MessageBox]::Show("The feature `Windows Sandbox` is not installed !!!")
        EXIT
    }
}

# Function to check if the Sandbox folder exists
function CheckFor-SandboxFolder {
    if ( [string]::IsNullOrEmpty($Sandbox_Folder) ) {
        return
    }
    if (-not (Test-Path -Path $Sandbox_Folder) ) {
        [System.Windows.Forms.MessageBox]::Show("Can not find the folder $Sandbox_Folder")
        EXIT
    }
}

function Copy-Sources {
    try {
        Copy-Item -Path $Sources -Destination $env:ProgramData -Force -Recurse | Out-Null
        Write-LogMessage -Message_Type "SUCCESS" -Message "Sources have been copied in $env:ProgramData\Run_in_Sandbox"
    } catch {
        Write-LogMessage -Message_Type "ERROR" -Message "Sources have not been copied in $env:ProgramData\Run_in_Sandbox"
        EXIT
    }
    
    if (-not (Test-Path -Path "$env:ProgramData\Run_in_Sandbox\RunInSandbox.ps1") ) {
        Write-LogMessage -Message_Type "ERROR" -Message "File RunInSandbox.ps1 is missing"
        [System.Windows.Forms.MessageBox]::Show("File RunInSandbox.ps1 is missing !!!")
        EXIT
    }
}

function Unblock-Sources {
    $Sources_Unblocked = $False
    try {
        Get-ChildItem -Path $Run_in_Sandbox_Folder -Recurse | Unblock-File
        Write-LogMessage -Message_Type "SUCCESS" -Message "Sources files have been unblocked"
        $Sources_Unblocked = $True
    } catch {
        Write-LogMessage -Message_Type "ERROR" -Message "Sources files have not been unblocked"
        EXIT
    }

    if ($Sources_Unblocked -ne $True) {
        Write-LogMessage -Message_Type "ERROR" -Message "Source files could not be unblocked"
        [System.Windows.Forms.MessageBox]::Show("Source files could not be unblocked")
        EXIT
    }
}

function Create-Checkpoint {
    if (-not $NoCheckpoint) {
        $SystemRestoreEnabled = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore" -Name "RPSessionInterval").RPSessionInterval
        if ($SystemRestoreEnabled -eq 0) {
            Write-LogMessage -Message_Type "WARNING" -Message "System Restore feature is disabled. Enable this to create a System restore point"
        } else {
            $Checkpoint_Command = '-Command Checkpoint-Computer -Description "Windows_Sandbox_Context_menus" -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop'
            $ReturnValue = Start-Process -FilePath "C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe" -ArgumentList $Checkpoint_Command -Wait -PassThru -WindowStyle Minimized
            if ($ReturnValue.ExitCode -eq 0) {
                Write-LogMessage -Message_Type "SUCCESS" -Message "Creation of restore point `"Add Windows Sandbox Context menus`""
            } else {
                Write-LogMessage -Message_Type "ERROR" -Message "Creation of restore point `"Add Windows Sandbox Context menus`" failed."
                Write-LogMessage -Message_Type "ERROR" -Message "Press any button to continue anyway."
                Read-Host
            }
        } 
    }
}