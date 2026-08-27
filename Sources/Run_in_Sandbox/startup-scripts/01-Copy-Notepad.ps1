# Runs inside Windows Sandbox. Assumes the host mapped ProgramData\Run_in_Sandbox to C:\Run_in_Sandbox in the Sandbox [1].
#
# The host maps its own Notepad++ installation read only to "C:\Program Files\Notepad++" when it has
# one. Only if there is no Notepad++ on the host, the classic notepad.exe is staged in NotepadPayload
# and gets copied into the sandbox by this script.
# The file name is kept as it is on purpose: a renamed script would leave the old one behind in
# C:\ProgramData\Run_in_Sandbox\startup-scripts of existing installations, and it would keep running.

$srcRoot = "C:\Run_in_Sandbox\NotepadPayload"
$srcSys  = Join-Path $srcRoot "System32"
$dstSys  = "C:\Windows\System32"
$dst  = "C:\Windows"

$notepadPlusPlusPath = "C:\Program Files\Notepad++\notepad++.exe"

function Register-Editor {
    param (
        [Parameter(Mandatory=$true)] [string]$EditorPath,
        [Parameter(Mandatory=$true)] [string]$EditorName
    )

    # "Edit with <editor>" for every file type
    reg add "HKEY_CLASSES_ROOT\*\shell\Edit with $EditorName" /f
    reg add "HKEY_CLASSES_ROOT\*\shell\Edit with $EditorName" /v "Icon" /t REG_SZ /d "$EditorPath,0" /f
    reg add "HKEY_CLASSES_ROOT\*\shell\Edit with $EditorName\command" /ve /d "`"$EditorPath`" `"%1`"" /f

    # "Open <editor>" on the background of a folder
    reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\$EditorName" /ve /d "Open $EditorName" /f
    reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\$EditorName" /v "Icon" /t REG_SZ /d "$EditorPath,0" /f
    reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\$EditorName\command" /ve /d "`"$EditorPath`"" /f

    # Make it the default for .txt
    cmd /c assoc .txt=txtfile
    If ( -not (Test-Path 'HKLM:\SOFTWARE\Classes\txtfile\shell\open\command') ) {
        New-Item -Path 'HKLM:\SOFTWARE\Classes\txtfile\shell\open\command' -Force | Out-Null
    }
    cmd /c ftype txtfile=`"$EditorPath`" "%1"
}

# 1) Notepad++ mapped by the host - nothing to copy, it can be used right away
if (Test-Path -LiteralPath $notepadPlusPlusPath) {
    Write-Host "[Copy-Notepad] Using the Notepad++ installation mapped by the host."
    Register-Editor -EditorPath $notepadPlusPlusPath -EditorName "Notepad++"
    exit 0
}

$hadError = $false
$notepadStaged = $false

# 2) Copy notepad.exe if present
try {
    $srcExe = Join-Path $srcSys "notepad.exe"
    $notepadPathSys = Join-Path $dstSys "notepad.exe"
    $notepadPath = Join-Path $dst "notepad.exe"
    if (Test-Path -LiteralPath $srcExe) {
        Copy-Item -LiteralPath $srcExe -Destination $notepadPath -Force
        Copy-Item -LiteralPath $srcExe -Destination $notepadPathSys -Force
        $notepadStaged = $true
        Write-Host "[Copy-Notepad] notepad.exe copied."
    } else {
        Write-Warning "[Copy-Notepad] Source not found: $srcExe"
    }
} catch {
    Write-Warning "[Copy-Notepad] Failed to copy notepad.exe: $($_.Exception.Message)"
    $hadError = $true
}

# 3) Copy the MUI file from the language subfolder(s) (e.g., en-US, de-DE)
try {
    $langDirs = Get-ChildItem -LiteralPath $srcSys -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match '^[a-z]{2}-[A-Z]{2}$' }

    foreach ($langDir in $langDirs) {
        $srcMui = Join-Path $langDir.FullName "notepad.exe.mui"
        if (Test-Path -LiteralPath $srcMui) {
            $dstLangDir = Join-Path $dst $langDir.Name
            $dstLangDirSys = Join-Path $dstSys $langDir.Name
            if (-not (Test-Path -LiteralPath $dstLangDir)) {
                New-Item -ItemType Directory -Path $dstLangDir -Force | Out-Null
            }
            if (-not (Test-Path -LiteralPath $dstLangDirSys)) {
                New-Item -ItemType Directory -Path $dstLangDirSys -Force | Out-Null
            }
            $dstMui = Join-Path $dstLangDir "notepad.exe.mui"
            $dstMuiSys = Join-Path $dstLangDirSys "notepad.exe.mui"
            if (-not (Test-Path -LiteralPath $dstMui) ) {
                Copy-Item -LiteralPath $srcMui -Destination $dstMui -Force
                Write-Host "[Copy-Notepad] MUI copied for $($langDir.Name)."
            }
            if (-not (Test-Path -LiteralPath $dstMuiSys) ) {
                Copy-Item -LiteralPath $srcMui -Destination $dstMuiSys -Force
                Write-Host "[Copy-Notepad] MUI copied for $($langDir.Name)."
            }
        } else {
            Write-Warning "[Copy-Notepad] MUI not found in: $($langDir.FullName)"
        }
    }
} catch {
    Write-Warning "[Copy-Notepad] Failed to copy MUI: $($_.Exception.Message)"
    $hadError = $true
}

if ($hadError) {
    exit 1
} elseif (-not $notepadStaged) {
    # Without an editor the entries below would point to a file that does not exist
    Write-Warning "[Copy-Notepad] Neither Notepad++ nor a notepad.exe is available, skipping the file associations."
    exit 0
} else {
    Register-Editor -EditorPath $notepadPath -EditorName "Notepad"

    # Restart Explorer so changes take effect
    # Get-Process explorer | Stop-Process -Force
    # Open an explorer window to the host-shared folder on first launch
    # Start-Process explorer.exe C:\Users\WDAGUtilityAccount\Desktop\HostShared

    exit 0
}
