[CmdletBinding()]
param(
    [string]$RepositoryUrl = "https://github.com/13349811148/skill.git",
    [string]$Branch = "main",
    [string]$CodexHome,
    [string]$WorkBuddyHome,
    [string]$CodeBuddyHome
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) ("shared-skill-bootstrap-" + [Guid]::NewGuid().ToString("N"))
$archivePath = Join-Path $temporaryRoot "repository.zip"
$extractRoot = Join-Path $temporaryRoot "repository"
New-Item -ItemType Directory -Force -Path $temporaryRoot, $extractRoot | Out-Null

try {
    $match = [regex]::Match($RepositoryUrl.Trim(), '^(?:https?://github\.com/)([^/]+)/([^/]+?)(?:\.git)?/?$')
    if (-not $match.Success) {
        throw "Only public GitHub repository URLs are supported: $RepositoryUrl"
    }
    $owner = [Uri]::EscapeDataString($match.Groups[1].Value)
    $repository = [Uri]::EscapeDataString($match.Groups[2].Value)
    $escapedBranch = [Uri]::EscapeDataString($Branch)
    $archiveUrl = "https://codeload.github.com/$owner/$repository/zip/refs/heads/$escapedBranch"

    Invoke-WebRequest -Uri $archiveUrl -OutFile $archivePath -UseBasicParsing -Headers @{
        "User-Agent" = "shared-skill-bootstrap"
    }
    Expand-Archive -LiteralPath $archivePath -DestinationPath $extractRoot -Force

    $installer = Get-ChildItem -Recurse -File -LiteralPath $extractRoot -Filter "install-skill-sync.ps1" |
        Select-Object -First 1 -ExpandProperty FullName
    if ([string]::IsNullOrWhiteSpace($installer)) {
        throw "The downloaded repository does not contain install-skill-sync.ps1."
    }

    $installerArguments = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", $installer,
        "-RepositoryUrl", $RepositoryUrl,
        "-Branch", $Branch
    )
    if ($PSBoundParameters.ContainsKey("CodexHome")) {
        $installerArguments += @("-CodexHome", $CodexHome)
    }
    if ($PSBoundParameters.ContainsKey("WorkBuddyHome")) {
        $installerArguments += @("-WorkBuddyHome", $WorkBuddyHome)
    }
    if ($PSBoundParameters.ContainsKey("CodeBuddyHome")) {
        $installerArguments += @("-CodeBuddyHome", $CodeBuddyHome)
    }

    & powershell.exe @installerArguments
    if ($LASTEXITCODE -ne 0) {
        throw "Skill synchronization installer failed with exit code $LASTEXITCODE."
    }
} finally {
    if (Test-Path -LiteralPath $temporaryRoot) {
        Remove-Item -Recurse -Force -LiteralPath $temporaryRoot
    }
}

