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
$Exported_Keys = [System.Collections.ArrayList]@()

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
        [string] $Type
    )
    
    # Every key only has to be exported once, before it is modified for the first time
    if ($Exported_Keys -contains $Reg_Path) {
        return
    }
    [void]$Exported_Keys.Add($Reg_Path)
    
    if (-not (Test-Path $Backup_Folder) ) {
        New-Item -ItemType Directory -Path $Backup_Folder -Force | Out-Null
    }
    
    Write-LogMessage -Message_Type "INFO" -Message "Exporting registry keys"
    
    # The registry path is part of the file name, but it contains characters that are not
    # allowed there. Without replacing them "SystemFileAssociations\.html" would point into a
    # folder that does not exist and "reg export" fails. The full path (including the hive) is
    # used so that backups of the same key name in HKCR and HKU do not overwrite each other
    $Backup_Name = "Backup_" + $Type + "_" + $Reg_Path
    $Backup_Name = $Backup_Name -replace '[\\/:*?"<>|]', '_'
    $Backup_Path = Join-Path -Path $Backup_Folder -ChildPath "$Backup_Name.reg"
    
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
    
    Export-RegConfig -Reg_Path $($Base_Registry_Key.Split("::")[-1]) -Type $Type -ErrorAction Continue
    
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

# Function to find 7-Zip installation on host system
function Find-Host7Zip {
    # Try common installation paths
    $CommonPaths = @(
        "C:\Program Files\7-Zip\7z.exe",
        "C:\Program Files (x86)\7-Zip\7z.exe",
        "${env:ProgramFiles}\7-Zip\7z.exe",
        "${env:ProgramFiles(x86)}\7-Zip\7z.exe"
    )
    
    foreach ($Path in $CommonPaths) {
        if (Test-Path $Path) {
            return $Path
        }
    }
    
    # Check registry for installation path
    try {
        $RegPath = Get-ItemProperty -Path "HKLM:\SOFTWARE\7-Zip" -Name "Path" -ErrorAction SilentlyContinue
        if ($RegPath -and (Test-Path "$($RegPath.Path)\7z.exe")) {
            return "$($RegPath.Path)\7z.exe"
        }
    } catch {}
    
    # Check PATH environment variable
    try {
        $7zInPath = Get-Command "7z.exe" -ErrorAction SilentlyContinue
        if ($7zInPath) {
            return $7zInPath.Source
        }
    } catch {}
    
    return $null
}

# Function to find a Notepad++ installation on the host system
function Find-HostNotepadPlusPlus {
    # Try common installation paths
    $CommonPaths = @(
        "${env:ProgramFiles}\Notepad++\notepad++.exe",
        "${env:ProgramFiles(x86)}\Notepad++\notepad++.exe",
        "C:\Program Files\Notepad++\notepad++.exe",
        "C:\Program Files (x86)\Notepad++\notepad++.exe"
    )

    foreach ($Path in $CommonPaths) {
        if ( (-not [string]::IsNullOrEmpty($Path)) -and (Test-Path -LiteralPath $Path) ) {
            return $Path
        }
    }

    # Check the registry, this also finds installations on another drive
    $RegistryPaths = @(
        "HKLM:\SOFTWARE\Notepad++",
        "HKLM:\SOFTWARE\WOW6432Node\Notepad++"
    )

    foreach ($RegistryPath in $RegistryPaths) {
        try {
            $Install_Folder = (Get-ItemProperty -Path $RegistryPath -ErrorAction SilentlyContinue)."(default)"
            if ( (-not [string]::IsNullOrEmpty($Install_Folder)) -and (Test-Path -LiteralPath "$Install_Folder\notepad++.exe") ) {
                return "$Install_Folder\notepad++.exe"
            }
        } catch {}
    }

    # Check the uninstall entry
    $UninstallPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Notepad++",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Notepad++"
    )

    foreach ($UninstallPath in $UninstallPaths) {
        try {
            $Install_Folder = (Get-ItemProperty -Path $UninstallPath -Name "InstallLocation" -ErrorAction SilentlyContinue).InstallLocation
            if ( (-not [string]::IsNullOrEmpty($Install_Folder)) -and (Test-Path -LiteralPath "$Install_Folder\notepad++.exe") ) {
                return "$Install_Folder\notepad++.exe"
            }
        } catch {}
    }

    # Check PATH environment variable, this also covers portable installations added to PATH
    try {
        $NotepadPlusPlusInPath = Get-Command "notepad++.exe" -ErrorAction SilentlyContinue
        if ($NotepadPlusPlusInPath) {
            return $NotepadPlusPlusInPath.Source
        }
    } catch {}

    return $null
}

# Function to get latest 7-Zip download URL from GitHub releases
function Get-Latest7ZipDownloadUrl {
    try {
        $ApiUrl = "https://api.github.com/repos/ip7z/7zip/releases/latest"
        $Response = Invoke-RestMethod -Uri $ApiUrl -UseBasicParsing
        
        # Look for x64 MSI installer first, fallback to x86 if needed
        $Asset = $Response.assets | Where-Object { $_.name -like "*-x64.msi" -and $_.name -notlike "*extra*" }
        
        if (-not $Asset) {
            # Fallback to x86 MSI if x64 not available
            $Asset = $Response.assets | Where-Object { $_.name -like "*.msi" -and $_.name -notlike "*extra*" -and $_.name -notlike "*x64*" }
        }
        
        if ($Asset) {
            return $Asset.browser_download_url
        }
    } catch {
        Write-LogMessage -Message_Type "WARNING" -Message "Failed to get latest 7-Zip version from GitHub: $($_.Exception.Message)"
    }
    
    return $null
}

# Function to check if cached 7-Zip installer should be updated
function Test-7ZipCacheAge {
    $TempFolder = "$Run_in_Sandbox_Folder\temp"
    $CachedInstaller = "$TempFolder\7zSetup.msi"
    $VersionFile = "$TempFolder\7zVersion.txt"
    
    # If no cached installer exists, we need to download
    if (-not (Test-Path $CachedInstaller)) {
        return $true
    }
    
    # Check if cache is older than 7 days
    $CacheAge = (Get-Date) - (Get-Item $CachedInstaller).LastWriteTime
    if ($CacheAge.Days -gt 7) {
        Write-LogMessage -Message_Type "INFO" -Message "Cached 7-Zip installer is $($CacheAge.Days) days old, checking for updates"
        return $true
    }
    
    return $false
}

# Function to download and cache latest 7-Zip installer
function Update-7ZipCache {
    param(
        [switch]$Force
    )
    
    $TempFolder = "$Run_in_Sandbox_Folder\temp"
    $CachedInstaller = "$TempFolder\7zSetup.msi"
    $VersionFile = "$TempFolder\7zVersion.txt"
    
    # Check if we need to update (unless forced)
    if (-not $Force -and -not (Test-7ZipCacheAge)) {
        Write-LogMessage -Message_Type "INFO" -Message "Cached 7-Zip installer is recent, skipping update"
        return $true
    }
    
    # Ensure temp folder exists
    if (-not (Test-Path $TempFolder)) {
        New-Item -ItemType Directory -Path $TempFolder -Force | Out-Null
    }
    
    # Get latest download URL
    $DownloadUrl = Get-Latest7ZipDownloadUrl
    if (-not $DownloadUrl) {
        Write-LogMessage -Message_Type "ERROR" -Message "Could not determine latest 7-Zip download URL"
        return $false
    }
    
    try {
        Write-LogMessage -Message_Type "INFO" -Message "Downloading latest 7-Zip installer from: $DownloadUrl"
        Invoke-WebRequest -Uri $DownloadUrl -OutFile $CachedInstaller -UseBasicParsing
        
        # Save download timestamp and URL for tracking
        @{
            Downloaded = (Get-Date).ToString()
            Url = $DownloadUrl
        } | ConvertTo-Json | Set-Content $VersionFile
        
        Write-LogMessage -Message_Type "SUCCESS" -Message "7-Zip installer cached successfully"
        return $true
    } catch {
        Write-LogMessage -Message_Type "ERROR" -Message "Failed to download 7-Zip installer: $($_.Exception.Message)"
        return $false
    }
}

# Function to ensure 7-Zip cache is available and current
function Ensure-7ZipCache {
    $TempFolder = "$Run_in_Sandbox_Folder\temp"
    $CachedInstaller = "$TempFolder\7zSetup.msi"
    
    # Try to update if needed (network available)
    try {
        if (Test-7ZipCacheAge) {
            Update-7ZipCache
        }
    } catch {
        Write-LogMessage -Message_Type "WARNING" -Message "Could not check for 7-Zip updates, using cached version if available"
    }
    
    # Return whether we have a usable cached installer
    return (Test-Path $CachedInstaller)
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
function Test-ForAdmin {
    $Run_As_Admin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $Run_As_Admin) {
        Write-LogMessage -Message_Type "ERROR" -Message "The script has not been launched with admin rights"
        [System.Windows.Forms.MessageBox]::Show("Please run the tool with admin rights :-)")
        EXIT
    }
    Write-LogMessage -Message_Type "INFO" -Message "The script has been launched with admin rights"
}

# Function to check for source files
function Test-ForSources {
    if (-not (Test-Path -Path $Sources)) {
        Write-LogMessage -Message_Type "ERROR" -Message "Sources folder is missing"
        [System.Windows.Forms.MessageBox]::Show("It seems you haven´t downloaded all the folder structure.`nThe folder `"Sources`" is missing !!!")
        EXIT
    }
    Write-LogMessage -Message_Type "SUCCESS" -Message "The sources folder exists"
    
    $Check_Sources_Files_Count = (Get-ChildItem -Path "$Current_Folder\Sources\Run_in_Sandbox" -Recurse).count
    if ($Check_Sources_Files_Count -lt 25) {  # Reduced from 40 to 26 (removed 14 bundled 7zip files)
        Write-LogMessage -Message_Type "ERROR" -Message "Some contents are missing"
        [System.Windows.Forms.MessageBox]::Show("It seems you haven´t downloaded all the folder structure !!!")
        EXIT
    }
}

# Function to check if the Windows Sandbox feature is installed
function Test-ForSandbox {
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
function Test-ForSandboxFolder {
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
    
    # Copy CommonFunctions.ps1 to the installation directory so RunInSandbox.ps1 can load it
    try {
        Copy-Item -Path "$Current_Folder\CommonFunctions.ps1" -Destination "$env:ProgramData\Run_in_Sandbox\" -Force | Out-Null
        Write-LogMessage -Message_Type "SUCCESS" -Message "CommonFunctions.ps1 copied to installation directory"
    } catch {
        Write-LogMessage -Message_Type "ERROR" -Message "Failed to copy CommonFunctions.ps1 to installation directory"
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

function New-Checkpoint {
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