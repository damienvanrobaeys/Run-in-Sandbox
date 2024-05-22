# Function to restart the script with admin rights
function Restart-ScriptWithAdmin {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Start-Process powershell.exe "-NoExit -NoProfile -ExecutionPolicy Bypass -Command `"(Invoke-webrequest -URI `"https://raw.githubusercontent.com/Joly0/Run-in-Sandbox/master/Install_Run-in-Sandbox.ps1`").Content | Invoke-Expression`"" -Verb RunAs
        exit
    }
}
# Restart the script with admin rights if not already running as admin
Restart-ScriptWithAdmin

# Define the URL and file paths
$zipUrl = "https://github.com/Joly0/Run-in-Sandbox/archive/refs/heads/master.zip"
$tempPath = [System.IO.Path]::GetTempPath()
$zipPath = Join-Path -Path $tempPath -ChildPath "master.zip"
$extractPath = Join-Path -Path $tempPath -ChildPath "Run-in-Sandbox-master"


# Remove existing extracted folder if it exists
if (Test-Path $extractPath) {
    try {
        Write-Host "Removing existing extracted folder..."
        Remove-Item -Path $extractPath -Recurse -Force
        Write-Host "Existing extracted folder removed."
    } catch {
        Write-Error "Failed to remove existing extracted folder: $_"
        exit 1
    }
}

# Download the zip file
try {
    Write-Host "Downloading zip file..."
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing
    $ProgressPreference = 'Continue'
    Write-Host "Download completed."
} catch {
    Write-Error "Failed to download the zip file: $_"
    exit 1
}

# Extract the zip file
try {
    Write-Host "Extracting zip file..."
    $ProgressPreference = 'SilentlyContinue'
    Expand-Archive -Path $zipPath -DestinationPath $tempPath
    $ProgressPreference = 'Continue'
    Write-Host "Extraction completed."
} catch {
    Write-Error "Failed to extract the zip file: $_"
    exit 1
}

# Remove the zip file
try {
    Write-Host "Removing zip file..."
    Remove-Item -Path $zipPath
    Write-Host "Zip file removed."
} catch {
    Write-Error "Failed to remove the zip file: $_"
    exit 1
}

# Construct the path to the add_structure.ps1 script
$addStructureScript = Join-Path -Path $extractPath -ChildPath "Add_Structure.ps1"

# Execute the add_structure.ps1 script with the "-NoCheckpoint" parameter if it was provided
try {
    Write-Host "Executing Add_Structure.ps1 script..."
    if ($NoCheckpoint) {
        & $addStructureScript -NoCheckpoint
    } else {
        & $addStructureScript
    }
    Write-Host "Script execution completed."
} catch {
    Write-Error "Failed to execute add_structure.ps1: $_"
    exit 1
}

Read-Host "Installation finished. Press Enter to exit."