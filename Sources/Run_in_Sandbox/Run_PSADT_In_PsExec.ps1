Param (
    [Parameter(Mandatory = $False)]
    [string]$WorkDir
)

if (-not ($WorkDirk) ) {
    $WorkDir = Read-Host "Enter the path for workdir"
}

$PSexec = "c:\pstools\PSexec.exe"
$ServiceUI = "ServiceUI.exe -Process:explorer.exe Deploy-Application.exe -DeploymentType Install -DeployMode Interactive"

$command = "$workdir\$ServiceUI"

$cmd = "$psexec -w `"$workdir`" -si -accepteula $command"

& { Invoke-Expression $cmd }