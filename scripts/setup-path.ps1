[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$PaseoBin = "$env:LOCALAPPDATA\Programs\Paseo\resources\bin"
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$projectBin = Join-Path $projectRoot 'bin'
$requestedDirectories = @($projectBin, $PaseoBin)
$userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
$userDirectories = @($userPath -split ';' | Where-Object { $_.Trim() })
$addedDirectories = @()

foreach ($directory in $requestedDirectories) {
    $absoluteDirectory = (Resolve-Path -LiteralPath $directory).Path.TrimEnd('\')
    $alreadyPresent = $userDirectories | Where-Object {
        [Environment]::ExpandEnvironmentVariables($_).TrimEnd('\') -ieq $absoluteDirectory
    }
    if (-not $alreadyPresent) {
        $userDirectories += $absoluteDirectory
        $addedDirectories += $absoluteDirectory
    }
}

if ($addedDirectories.Count -eq 0) {
    Write-Output 'User PATH already contains the Office Viewer and Paseo directories.'
    return
}

if ($PSCmdlet.ShouldProcess('Current user PATH', "Add $($addedDirectories -join ', ')")) {
    [Environment]::SetEnvironmentVariable('Path', ($userDirectories -join ';'), 'User')
    $env:Path = $env:Path.TrimEnd(';') + ';' + ($addedDirectories -join ';')
    Write-Output "Added to user PATH: $($addedDirectories -join ', ')"
    Write-Output 'Open a new terminal to use the commands. Restart an existing terminal app if needed.'
}
