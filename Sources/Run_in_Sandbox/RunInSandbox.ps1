param (
    [Parameter(Mandatory=$true)] [String]$Type,
    [Parameter(Mandatory=$true)] [String]$ScriptPath
)

# ------------------------------------------------------------------------------------
# Logging and error handling
# The context menu starts this script with "powershell.exe -WindowStyle Hidden", so any
# terminating error would kill the script without the user noticing anything at all.
# Everything is written to a log file and unexpected errors are shown in a message box.
# ------------------------------------------------------------------------------------
$Run_Log_File = Join-Path -Path $env:TEMP -ChildPath "RunInSandbox.log"

function Write-RunLog {
    param (
        [string]$Message,
        [string]$Message_Type = "INFO"
    )

    $MyDate = "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
    Add-Content -Path $Run_Log_File -Value "$MyDate - $Message_Type : $Message" -ErrorAction SilentlyContinue
}

function Show-RunError {
    param (
        [string]$Message
    )

    Write-RunLog -Message_Type "ERROR" -Message $Message
    try {
        [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
        [System.Windows.Forms.MessageBox]::Show("$Message`n`nMore details can be found in `"$Run_Log_File`"", "Run in Sandbox", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    } catch {}
}

trap {
    Show-RunError -Message "Run in Sandbox stopped with an error: $($_.Exception.Message)"
    Write-RunLog -Message_Type "ERROR" -Message "$($_.ScriptStackTrace)"
    exit 1
}

Write-RunLog -Message_Type "INFO" -Message "Started for type `"$Type`" and path `"$ScriptPath`""

#Start-Transcript -Path $(Join-Path -Path $([System.Environment]::GetEnvironmentVariables('Machine').TEMP) -ChildPath "RunInSandbox.log")

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
    $DirectoryName = (Get-Item $ScriptPath).DirectoryName
    # Has to be taken from the resolved path, $ScriptPath is wildcard escaped and would
    # add backticks to folder names containing "[" or "]", which then do not exist in the sandbox
    $FolderPath = Split-Path $DirectoryName -Leaf
    $FileName = (Get-Item $ScriptPath).BaseName
    $Full_FileName = (Get-Item $ScriptPath).Name
}

$Sandbox_Desktop_Path = "C:\Users\WDAGUtilityAccount\Desktop"
$Sandbox_Shared_Path = "$Sandbox_Desktop_Path\$FolderPath"

$Sandbox_Root_Path = "C:\Run_in_Sandbox"
$Full_Startup_Path = "$Sandbox_Shared_Path\$Full_FileName"
$Full_Startup_Path_Quoted = """$Full_Startup_Path"""

$Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"

# Load common functions
. "$Run_in_Sandbox_Folder\CommonFunctions.ps1"

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

function Enable-StartupScripts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$OriginalCommand,  # whatever you would have passed as -Command_to_Run
        [string]$StartupScriptFolderName = "startup-scripts"
    )

    $StartupScriptsFolder = Join-Path $Run_in_Sandbox_Folder $StartupScriptFolderName
    New-Item -ItemType Directory -Path $StartupScriptsFolder -Force | Out-Null
    
    $origCmdFile = Join-Path $StartupScriptsFolder "OriginalCommand.txt"
    if ($OriginalCommand -ne "") {
       # Write the original command into a file
        try {
            Set-Content -LiteralPath $origCmdFile -Value $OriginalCommand -Encoding UTF8 -Force -ErrorAction Stop
        } catch {
            # Without this file nothing would happen inside the sandbox, so do not continue silently
            throw "The command for the sandbox could not be written to `"$origCmdFile`": $($_.Exception.Message)"
        }
    } else {
        # Nothing to run, make sure the command of a previous run is not executed again
        if (Test-Path -LiteralPath $origCmdFile) {
            Remove-Item -LiteralPath $origCmdFile -Force -ErrorAction SilentlyContinue
            if (Test-Path -LiteralPath $origCmdFile) {
                Write-RunLog -Message_Type "WARNING" -Message "Could not delete `"$origCmdFile`", the command of the previous run will be executed again"
            }
        }
    }

    # Orchestrator that runs NN-*.ps1 in lexicographic order, then runs the original command
    $orchestrator = @'
param(
    [string]$ScriptsPath = "C:\Run_in_Sandbox\startup-scripts",
    # Can include this switch when running from the .wsb file to indicate it's the first launch of the sandbox
    # Useful if re-running this script within the sandbox as a test, but don't want certain parts to run again
    [switch]$launchingSandbox
)

# ------ Check that we're running in the Windows Sandbox ------
# This script is intended to be run from within the Windows Sandbox. We'll do a rudamentary check for if the current user is named "WDAGUtilityAccount"
if ($env:USERNAME -ne "WDAGUtilityAccount") {
    Write-host "`n`nERROR: This script is intended to be run from WITHIN the Windows Sandbox.`nIt appears you are running this from outside the sandbox.`n" -ForegroundColor Red
    Write-host "`nPress Enter to exit." -ForegroundColor Yellow
    Read-Host
    exit
}

Write-Host "[Orchestrator] Scripts path: $ScriptsPath"

# 1) Run ordered startup scripts: 00-*, 01-* ... 99-*
$pattern = '^\d{2}-.+\.ps1$'
$items = Get-ChildItem -LiteralPath $ScriptsPath -Filter *.ps1 -File -ErrorAction SilentlyContinue |
         Where-Object { $_.Name -match $pattern } |
         Sort-Object Name

foreach ($i in $items) {
    Write-Host "[Orchestrator] Running: $($i.Name)"
    try {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $i.FullName
        $rc = $LASTEXITCODE
        if ($rc -ne $null -and $rc -ne 0) {
            Write-Warning "[Orchestrator] Script $($i.Name) returned exit code $rc"
        }
    } catch {
        Write-Warning "[Orchestrator] Script $($i.Name) threw: $($_.Exception.Message)"
    }
}

# Restart Explorer so changes take effect
Write-Host "[Orchestrator] Restarting Explorer so changes take effect"
Get-Process explorer | Stop-Process -Force

# 2) Read and run the original command last
$origFile = Join-Path $ScriptsPath "OriginalCommand.txt"
if (Test-Path -LiteralPath $origFile) {
    $orig = Get-Content -LiteralPath $origFile -Raw
    Write-Host "[Orchestrator] Running original command..."
    # Run through cmd to support both cmd and PowerShell-style lines
    Start-process -Filepath "C:\Windows\SysWOW64\cmd.exe" -ArgumentList @('/c', '"' + $orig + '"') -WindowStyle Hidden
} else {
    Write-Warning "[Orchestrator] OriginalCommand.txt not found; nothing to run."
}
'@

    $orchestratorPath = Join-Path $StartupScriptsFolder "_orchestrator.ps1"
    # The context menu runs WITHOUT elevation, while this file was put there by the installer
    # (elevated), so overwriting it is denied for a standard user. It only needs to be written
    # when it is missing or outdated, and if that fails the installed one is just as good
    $Orchestrator_OnDisk = ""
    if (Test-Path -LiteralPath $orchestratorPath) {
        $Orchestrator_OnDisk = Get-Content -LiteralPath $orchestratorPath -Raw -ErrorAction SilentlyContinue
        if ($null -eq $Orchestrator_OnDisk) {
            $Orchestrator_OnDisk = ""
        }
    }

    if ($Orchestrator_OnDisk.TrimEnd() -ne $orchestrator.TrimEnd()) {
        try {
            Set-Content -LiteralPath $orchestratorPath -Value $orchestrator -Encoding UTF8 -Force -ErrorAction Stop
        } catch {
            if (-not (Test-Path -LiteralPath $orchestratorPath) ) {
                throw
            }
            Write-RunLog -Message_Type "WARNING" -Message "Could not update `"$orchestratorPath`" ($($_.Exception.Message)), the installed version is used instead"
        }
    }

    # Return the single Sandbox command that runs the orchestrator
    "C:\Run_in_Sandbox\ServiceUI.exe -Process:explorer.exe C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -sta -WindowStyle Normal -NoProfile -ExecutionPolicy Bypass -NoExit -File `"$Sandbox_Root_Path\$StartupScriptFolderName\_orchestrator.ps1`""
}

function Add-NotepadToSandbox {
    [CmdletBinding()]
    param(
        [string]$HostPayloadRoot = "C:\ProgramData\Run_in_Sandbox\NotepadPayload",
        [switch]$EnforceEnUsFallback # if set, will try en-US when preferred language is missing
    )

    # Resolve a single notepad.exe (prefer System32)
    $exeCandidates = Get-Command notepad.exe -ErrorAction Stop | Select-Object -ExpandProperty Source
    # The app execution alias in WindowsApps is a 0 byte reparse point, copying it is useless
    $exeCandidates = $exeCandidates | Where-Object { (Test-Path -LiteralPath $_) -and ((Get-Item -LiteralPath $_).Length -gt 0) }
    $exePath = ($exeCandidates | Where-Object { $_ -match '\\Windows\\System32\\' } | Select-Object -First 1)
    if (-not $exePath) { $exePath = $exeCandidates | Select-Object -First 1 }
    if (-not $exePath) {
        throw "No usable notepad.exe found, the classic Notepad seems to be removed from this system."
    }

    $exeDir  = Split-Path $exePath -Parent
    $exeName = Split-Path $exePath -Leaf

    # Build candidate language list
    $candidates = @()
    try { $candidates += (Get-UICulture).Name } catch {}
    try { $candidates += (Get-WinSystemLocale).Name } catch {}
    try {
        $candidates += Get-ChildItem "HKLM:\SYSTEM\CurrentControlSet\Control\MUI\UILanguages" |
                       Select-Object -ExpandProperty PSChildName
    } catch {}
    $candidates = $candidates | Where-Object { $_ } | Select-Object -Unique

    # Probe possible MUI locations
    $dirs = @(
        $exeDir,
        (Join-Path $env:WINDIR 'System32'),
        (Join-Path $env:WINDIR 'SysWOW64'),
        $env:WINDIR
    ) | Select-Object -Unique

    $muiPath = $null; $resolvedLang = $null
    foreach ($lang in $candidates) {
        foreach ($dir in $dirs) {
            $p = Join-Path (Join-Path $dir $lang) "$exeName.mui"
            if (Test-Path -LiteralPath $p) { $muiPath = $p; $resolvedLang = $lang; break }
        }
        if ($muiPath) { break }
    }

    if (-not $muiPath -and $EnforceEnUsFallback) {
        foreach ($dir in $dirs) {
            $fallback = Join-Path (Join-Path $dir 'en-US') "$exeName.mui"
            if (Test-Path -LiteralPath $fallback) { $muiPath = $fallback; $resolvedLang = 'en-US'; break }
        }
    }

    # Stage payload on host: System32\notepad.exe and System32\<lang>\notepad.exe.mui
    # The .mui only holds the localized strings, notepad also runs without it, so a missing
    # one must not cost us the whole payload
    $sys32Out = Join-Path $HostPayloadRoot "System32"
    New-Item -ItemType Directory -Path $sys32Out -Force -ErrorAction Stop | Out-Null
    Copy-Item -LiteralPath $exePath -Destination (Join-Path $sys32Out $exeName) -Force -ErrorAction Stop

    if ($muiPath) {
        $langOut = Join-Path $sys32Out $resolvedLang
        New-Item -ItemType Directory -Path $langOut -Force -ErrorAction Stop | Out-Null
        Copy-Item -LiteralPath $muiPath -Destination (Join-Path $langOut "$exeName.mui") -Force -ErrorAction Stop
    } else {
        Write-RunLog -Message_Type "WARNING" -Message "No notepad.exe.mui found for $exePath, notepad will be copied without its localized strings"
    }
}

function New-WSB {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [String]$Command_to_Run,
        [Array]$AdditionalMappedFolders = @()
    )
    
    # Give the sandbox an editor. Notepad++ from the host is preferred and only mounted
    # read only, the same way the host installation of 7-Zip is used. Only when there is no
    # Notepad++ the classic Notepad is staged as a fallback.
    # Both are optional, neither may stop the sandbox from starting.
    $NotepadPlusPlus_Path = Find-HostNotepadPlusPlus
    if ($NotepadPlusPlus_Path) {
        $NotepadPlusPlus_Folder = Split-Path $NotepadPlusPlus_Path -Parent
        # Do not map it twice if the file that is opened lies in the Notepad++ folder itself
        if ($NotepadPlusPlus_Folder -ne $DirectoryName) {
            $AdditionalMappedFolders += @{
                HostFolder = $NotepadPlusPlus_Folder
                SandboxFolder = "C:\Program Files\Notepad++"
                ReadOnly = "true"
            }
        }
        Write-RunLog -Message_Type "INFO" -Message "Using the Notepad++ installation of the host: $NotepadPlusPlus_Path"
    } else {
        # Add-NotepadToSandbox throws when notepad.exe or its .mui file cannot be found,
        # which is the case on systems where the classic Notepad has been removed
        try {
            Add-NotepadToSandbox -EnforceEnUsFallback | Out-Null
        } catch {
            Write-RunLog -Message_Type "WARNING" -Message "No Notepad++ found and Notepad could not be prepared either: $($_.Exception.Message)"
        }
    }
    
    New-Item $Sandbox_File_Path -type file -Force | Out-Null
    Add-Content -LiteralPath $Sandbox_File_Path -Value "<Configuration>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "    <VGpu>$Sandbox_VGpu</VGpu>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "    <Networking>$Sandbox_Networking</Networking>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "    <AudioInput>$Sandbox_AudioInput</AudioInput>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "    <VideoInput>$Sandbox_VideoInput</VideoInput>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "    <ProtectedClient>$Sandbox_ProtectedClient</ProtectedClient>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "    <PrinterRedirection>$Sandbox_PrinterRedirection</PrinterRedirection>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "    <ClipboardRedirection>$Sandbox_ClipboardRedirection</ClipboardRedirection>"
    if ( -not [string]::IsNullOrEmpty($Sandbox_MemoryInMB) ) {
        Add-Content -LiteralPath $Sandbox_File_Path -Value "    <MemoryInMB>$Sandbox_MemoryInMB</MemoryInMB>"
    }

    Add-Content $Sandbox_File_Path "    <MappedFolders>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "        <MappedFolder>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "            <HostFolder>C:\ProgramData\Run_in_Sandbox</HostFolder>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "            <SandboxFolder>C:\Run_in_Sandbox</SandboxFolder>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "            <ReadOnly>$Sandbox_ReadOnlyAccess</ReadOnly>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "        </MappedFolder>"

    if ($Type -eq "SDBApp") {
        $SDB_Full_Path = $ScriptPath
        # AppBundle_Install.ps1 reads the bundle as "App_Bundle.sdbapp", so it has to be
        # copied under that name, no matter how the file the user picked is called
        Copy-Item $ScriptPath "$Run_in_Sandbox_Folder\App_Bundle.sdbapp" -Force
        $Get_Apps_to_install = [xml](Get-Content $SDB_Full_Path)
        $Apps_to_install_path = $Get_Apps_to_install.Applications.Application.Path | Select-Object -Unique

        ForEach ($App_Path in $Apps_to_install_path) {
            if ( [string]::IsNullOrEmpty($App_Path) ) {
                continue
            }
            # Every application folder gets its own folder below C:\SBDApp. Mapping all of them
            # to C:\SBDApp would collide as soon as a bundle contains more than one folder and
            # AppBundle_Install.ps1 looks for the files in C:\SBDApp\<name of the host folder>
            $App_Folder_Name = Split-Path $App_Path -Leaf
            Get-ChildItem -Path $App_Path -Recurse | Unblock-File
            Add-Content -LiteralPath $Sandbox_File_Path -Value "        <MappedFolder>"
            Add-Content -LiteralPath $Sandbox_File_Path -Value "            <HostFolder>$App_Path</HostFolder>"
            Add-Content -LiteralPath $Sandbox_File_Path -Value "            <SandboxFolder>C:\SBDApp\$App_Folder_Name</SandboxFolder>"
            Add-Content -LiteralPath $Sandbox_File_Path -Value "            <ReadOnly>$Sandbox_ReadOnlyAccess</ReadOnly>"
            Add-Content -LiteralPath $Sandbox_File_Path -Value "        </MappedFolder>"
        }
    } else {
        Get-ChildItem -Path $DirectoryName -Recurse | Unblock-File
        Add-Content -LiteralPath $Sandbox_File_Path -Value "        <MappedFolder>"
        Add-Content -LiteralPath $Sandbox_File_Path -Value "            <HostFolder>$DirectoryName</HostFolder>"
        if ($Type -eq "IntuneWin") { Add-Content -LiteralPath $Sandbox_File_Path -Value "            <SandboxFolder>C:\IntuneWin</SandboxFolder>" }
        Add-Content -LiteralPath $Sandbox_File_Path -Value "            <ReadOnly>$Sandbox_ReadOnlyAccess</ReadOnly>"
        Add-Content -LiteralPath $Sandbox_File_Path -Value "        </MappedFolder>"
    }
    
    # Add any additional mapped folders
    foreach ($MappedFolder in $AdditionalMappedFolders) {
        Get-ChildItem -Path $($MappedFolder.HostFolder) -Recurse | Unblock-File
        Add-Content -LiteralPath $Sandbox_File_Path -Value "        <MappedFolder>"
        Add-Content -LiteralPath $Sandbox_File_Path -Value "            <HostFolder>$($MappedFolder.HostFolder)</HostFolder>"
        Add-Content -LiteralPath $Sandbox_File_Path -Value "            <SandboxFolder>$($MappedFolder.SandboxFolder)</SandboxFolder>"
        Add-Content -LiteralPath $Sandbox_File_Path -Value "            <ReadOnly>$($MappedFolder.ReadOnly)</ReadOnly>"
        Add-Content -LiteralPath $Sandbox_File_Path -Value "        </MappedFolder>"
    }
    Add-Content -LiteralPath $Sandbox_File_Path -Value "    </MappedFolders>"
    
    Add-Content -Path $Sandbox_File_Path  -Value "    <LogonCommand>"
    Add-Content -Path $Sandbox_File_Path  -Value "        <Command>$Command_to_Run</Command>"
    Add-Content -Path $Sandbox_File_Path  -Value "    </LogonCommand>"
    Add-Content -Path $Sandbox_File_Path  -Value "</Configuration>"
}

switch ($Type) {
    "7Z" {
        # Try to find 7-Zip on host system first
        $Host7ZipPath = Find-Host7Zip
        $AdditionalFolders = @()
        
        if ($Host7ZipPath) {
            # Mount the host 7-Zip installation into sandbox
            $Host7ZipFolder = Split-Path $Host7ZipPath -Parent
            
            $AdditionalFolders += @{
                HostFolder = $Host7ZipFolder
                SandboxFolder = "C:\Program Files\7-Zip"
                ReadOnly = "true"
            }
            
            $Script:Startup_Command = "`"C:\Program Files\7-Zip\7z.exe`" x $Full_Startup_Path_Quoted -y -oC:\Users\WDAGUtilityAccount\Desktop\Extracted_File"
            
            Write-LogMessage -Message_Type "INFO" -Message "Using host 7-Zip installation: $Host7ZipPath"
        }
        else {
            # No host installation found, ensure we have a cached installer
            if (-not (Ensure-7ZipCache)) {
                [System.Windows.Forms.MessageBox]::Show("Failed to download 7-Zip installer and no cached version available.`nPlease check your internet connection.")
                EXIT
            }
            
            $CachedInstaller = "$Sandbox_Root_Path\temp\7zSetup.msi"
            
            # Install 7-Zip in sandbox then extract
            $Script:Startup_Command = "$PSRun_Command `"Start-Process -FilePath 'msiexec.exe' -ArgumentList '/i \`"$CachedInstaller\`" /quiet' -Wait; Start-Process -FilePath 'C:\Program Files\7-Zip\7z.exe' -ArgumentList 'x $Full_Startup_Path_Quoted -y -oC:\Users\WDAGUtilityAccount\Desktop\Extracted_File' -Wait`""
            
            Write-LogMessage -Message_Type "INFO" -Message "Using cached 7-Zip installer: $CachedInstaller"
        }
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -AdditionalMappedFolders $AdditionalFolders
    }
    "CMD" {
        $Script:Startup_Command = $PSRun_Command + " " + "Start-Process $Full_Startup_Path_Quoted"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "EXE" {
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
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "Folder_On" {
        $Startup_Command = Enable-StartupScripts
        New-WSB -Command_to_Run $Startup_Command
    }
    "Folder_Inside" {
        $Startup_Command = Enable-StartupScripts
        New-WSB -Command_to_Run $Startup_Command
    }
    "HTML" {
        $Script:Startup_Command = $PSRun_Command + " " + "`"Invoke-Item -Path `'$Full_Startup_Path`'`""
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "URL" {
        $Script:Startup_Command = $PSRun_Command + " " + "Start-Process $Sandbox_Root_Path"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "Intunewin" {
        $Intunewin_Folder = "C:\IntuneWin\$FileName.intunewin"
        $Intunewin_Content_File = "$Run_in_Sandbox_Folder\Intunewin_Folder.txt"
        $Intunewin_Command_File = "$Run_in_Sandbox_Folder\Intunewin_Install_Command.txt"
        $Intunewin_Folder | Out-File $Intunewin_Content_File -Force -NoNewline

        #$Full_Startup_Path_UnQuoted = $Full_Startup_Path_Quoted.Replace('"', "")

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
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "ISO" {
        # Try to find 7-Zip on host system first
        $Host7ZipPath = Find-Host7Zip
        $AdditionalFolders = @()
        
        if ($Host7ZipPath) {
            # Mount the host 7-Zip installation into sandbox
            $Host7ZipFolder = Split-Path $Host7ZipPath -Parent
            
            $AdditionalFolders += @{
                HostFolder = $Host7ZipFolder
                SandboxFolder = "C:\Program Files\7-Zip"
                ReadOnly = "true"
            }
            
            $Script:Startup_Command = "`"C:\Program Files\7-Zip\7z.exe`" x $Full_Startup_Path_Quoted -y -oC:\Users\WDAGUtilityAccount\Desktop\Extracted_ISO"
            
            Write-LogMessage -Message_Type "INFO" -Message "Using host 7-Zip installation for ISO: $Host7ZipPath"
        }
        else {
            # No host installation found, ensure we have a cached installer
            if (-not (Ensure-7ZipCache)) {
                [System.Windows.Forms.MessageBox]::Show("Failed to download 7-Zip installer and no cached version available.`nPlease check your internet connection.")
                EXIT
            }
            
            $CachedInstaller = "$Run_in_Sandbox_Folder\temp\7zSetup.msi"
            
            # Install 7-Zip in sandbox then extract ISO
            $Script:Startup_Command = "$PSRun_Command `"Start-Process -FilePath 'msiexec.exe' -ArgumentList '/i \`"$CachedInstaller\`" /quiet' -Wait; Start-Process -FilePath 'C:\Program Files\7-Zip\7z.exe' -ArgumentList 'x $Full_Startup_Path_Quoted -y -oC:\Users\WDAGUtilityAccount\Desktop\Extracted_ISO' -Wait`""
            
            Write-LogMessage -Message_Type "INFO" -Message "Using cached 7-Zip installer for ISO: $CachedInstaller"
        }
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -AdditionalMappedFolders $AdditionalFolders
    }
    "MSI" {
        $Full_Startup_Path_UnQuoted = $Full_Startup_Path_Quoted.Replace('"', "")

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
                $Script:Startup_Command = "msiexec /i `"$Full_Startup_Path_UnQuoted`" " + $Switches_MSI
                $Form_MSI.close()
            })

        $Form_MSI.Add_Closing({
                $Script:Switches_MSI = $switches_for_exe.Text.ToString()
                $Script:Startup_Command = "msiexec /i `"$Full_Startup_Path_UnQuoted`" " + $Switches_MSI
            })

        $Form_MSI.ShowDialog() | Out-Null

        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "MSIX" {
        $Script:Startup_Command = $PSRun_Command + " " + "Add-AppPackage -Path $Full_Startup_Path_Quoted"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "PDF" {
        $Full_Startup_Path_Quoted = $Full_Startup_Path_Quoted.Replace('"', '')
        $Script:Startup_Command = $PSRun_Command + " " + "`"Invoke-Item -Path `'$Full_Startup_Path_Quoted`'`""
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "PPKG" {
        $Script:Startup_Command = $PSRun_Command + " " + "Install-ProvisioningPackage $Full_Startup_Path_Quoted -forceinstall -quietinstall"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "PS1Basic" {
        $Script:Startup_Command = $PSRun_File + " " + "$Full_Startup_Path_Quoted"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "PS1System" {
        $Script:Startup_Command = "$Sandbox_Root_Path\PsExec.exe \\localhost -nobanner -accepteula -s Powershell -ExecutionPolicy Bypass -File $Full_Startup_Path_Quoted"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "PS1Params" {
        $Full_Startup_Path_UnQuoted = $Full_Startup_Path_Quoted.Replace('"', "")

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
                $Script:Startup_Command = $PSRun_File + " " + "`"$Full_Startup_Path_UnQuoted`"" + " " + "$Paramaters"
                $Form_PS1.close()
            })

        $Form_PS1.Add_Closing({
                $Script:Paramaters = $parameters_to_add.Text.ToString()
                $Script:Startup_Command = $PSRun_File + " " + "`"$Full_Startup_Path_UnQuoted`"" + " " + "$Paramaters"
            })

        $Form_PS1.ShowDialog() | Out-Null
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "REG" {
        $Script:Startup_Command = "REG IMPORT $Full_Startup_Path_Quoted"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "SDBApp" {
        $AppBundle_Installer = "$Sandbox_Root_Path\AppBundle_Install.ps1"
        $Script:Startup_Command = $PSRun_File + " " + "$AppBundle_Installer"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "VBSBasic" {
        $Script:Startup_Command = "wscript.exe $Full_Startup_Path_Quoted"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "VBSParams" {
        $Full_Startup_Path_UnQuoted = $Full_Startup_Path_Quoted.Replace('"', '')

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
                $Script:Startup_Command = "wscript.exe `"$Full_Startup_Path_UnQuoted`" $Paramaters"
                $Form_VBS.close()
            })

        $Form_VBS.Add_Closing({
                $Script:Paramaters = $parameters_to_add.Text.ToString()
                $Script:Startup_Command = "wscript.exe `"$Full_Startup_Path_UnQuoted`" $Paramaters"
            })

        $Form_VBS.ShowDialog() | Out-Null

        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "ZIP" {
        $Script:Startup_Command = $PSRun_Command + " " + "`"Expand-Archive -LiteralPath '$Full_Startup_Path' -DestinationPath '$Sandbox_Desktop_Path\ZIP_extracted'`""
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    default {
        Show-RunError -Message "The type `"$Type`" is not supported, nothing can be started in the sandbox."
        exit 1
    }
}

if (-not (Test-Path -LiteralPath $Sandbox_File_Path) ) {
    Show-RunError -Message "No sandbox configuration file has been created for type `"$Type`", the sandbox cannot be started."
    exit 1
}

Write-RunLog -Message_Type "INFO" -Message "Starting the sandbox with `"$Sandbox_File_Path`""
Start-Process -FilePath $Sandbox_File_Path -Wait
do {
    Start-Sleep -Seconds 1
} while (Get-Process -Name "WindowsSandboxServer" -ErrorAction SilentlyContinue)

if ($WSB_Cleanup -eq $True) {
    # Most of these variables are only set by one of the types, for every other type they are
    # $null. Passing $null to -LiteralPath is a parameter binding error, and that one is
    # terminating and NOT suppressed by -ErrorAction SilentlyContinue, so empty entries have to
    # be skipped - otherwise the script dies here at the end of every single run
    $Files_to_Cleanup = @(
        $Sandbox_File_Path
        $Intunewin_Command_File
        $Intunewin_Content_File
        $EXE_Command_File
        "$Run_in_Sandbox_Folder\App_Bundle.sdbapp"
        "$Run_in_Sandbox_Folder\NotepadPayload"
        "$Run_in_Sandbox_Folder\startup-scripts\OriginalCommand.txt"
    )

    ForEach ($File_to_Cleanup in $Files_to_Cleanup) {
        if ( [string]::IsNullOrEmpty($File_to_Cleanup) ) {
            continue
        }
        Remove-Item -LiteralPath $File_to_Cleanup -Force -Recurse -ErrorAction SilentlyContinue
    }
}

Write-RunLog -Message_Type "INFO" -Message "Finished"