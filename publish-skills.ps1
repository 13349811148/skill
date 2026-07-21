[CmdletBinding()]
param(
    [string]$Message,
    [string]$CodexHome,
    [string]$SkillsDir,
    [switch]$NoPush
)

$ErrorActionPreference = "Stop"

$pythonCommand = Get-Command python3 -ErrorAction SilentlyContinue
$pythonPrefix = @()
if ($null -eq $pythonCommand) {
    $pythonCommand = Get-Command python -ErrorAction SilentlyContinue
}
if ($null -eq $pythonCommand) {
    $pythonCommand = Get-Command py -ErrorAction SilentlyContinue
    if ($null -ne $pythonCommand) {
        $pythonPrefix = @("-3")
    }
}
if ($null -eq $pythonCommand) {
    throw "Python 3 is required. Install Python 3 and run this command again."
}

$publisherArguments = @($pythonPrefix)
$publisherArguments += (Join-Path $PSScriptRoot "publish-skills.py")
if (-not [string]::IsNullOrWhiteSpace($Message)) {
    $publisherArguments += @("--message", $Message)
}
if (-not [string]::IsNullOrWhiteSpace($CodexHome)) {
    $publisherArguments += @("--client-home", $CodexHome)
}
if (-not [string]::IsNullOrWhiteSpace($SkillsDir)) {
    $publisherArguments += @("--skills-dir", $SkillsDir)
}
if ($NoPush) {
    $publisherArguments += "--no-push"
}

& $pythonCommand.Source @publisherArguments
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}
