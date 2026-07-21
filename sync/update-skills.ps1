[CmdletBinding()]
param(
    [string]$CodexHome,
    [string]$RepositoryUrl = "https://github.com/13349811148/skill.git",
    [string]$Branch = "main",
    [switch]$Force,
    [switch]$HookMode
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

function Resolve-CodexHome {
    param([string]$RequestedHome)

    if (-not [string]::IsNullOrWhiteSpace($RequestedHome)) {
        return [IO.Path]::GetFullPath($RequestedHome)
    }
    if ((Split-Path -Leaf $PSScriptRoot) -eq "scripts") {
        $skillDirectory = Split-Path -Parent $PSScriptRoot
        $skillsDirectory = Split-Path -Parent $skillDirectory
        if ((Split-Path -Leaf $skillsDirectory) -eq "skills") {
            return Split-Path -Parent $skillsDirectory
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($env:CODEX_HOME)) {
        return [IO.Path]::GetFullPath($env:CODEX_HOME)
    }
    return Join-Path $env:USERPROFILE ".codex"
}

function Write-HookJson {
    param(
        [string]$SystemMessage,
        [string]$AdditionalContext
    )

    $result = [ordered]@{ continue = $true }
    if (-not [string]::IsNullOrWhiteSpace($SystemMessage)) {
        $result.systemMessage = $SystemMessage
    }
    if (-not [string]::IsNullOrWhiteSpace($AdditionalContext)) {
        $result.hookSpecificOutput = [ordered]@{
            hookEventName = "SessionStart"
            additionalContext = $AdditionalContext
        }
    }
    Write-Output ($result | ConvertTo-Json -Depth 6 -Compress)
}

function Resolve-GithubRepository {
    param([string]$RepositoryUrl)

    $match = [regex]::Match($RepositoryUrl.Trim(), '^(?:https?://github\.com/|git@github\.com:)([^/]+)/([^/]+?)(?:\.git)?/?$')
    if (-not $match.Success) {
        throw "Only public GitHub repository URLs are supported: $RepositoryUrl"
    }
    return [pscustomobject]@{
        Owner = $match.Groups[1].Value
        Repository = $match.Groups[2].Value
    }
}

function Invoke-PublicDownload {
    param(
        [string]$Uri,
        [string]$OutFile
    )

    $parameters = @{
        Uri = $Uri
        UseBasicParsing = $true
        Headers = @{ "User-Agent" = "shared-skill-sync" }
    }
    if (-not [string]::IsNullOrWhiteSpace($OutFile)) {
        $parameters.OutFile = $OutFile
    }
    return Invoke-WebRequest @parameters
}

function Read-JsonFile {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Required file does not exist: $Path"
    }
    return Get-Content -Raw -LiteralPath $Path -Encoding UTF8 | ConvertFrom-Json
}

function Write-JsonFileAtomic {
    param(
        [string]$Path,
        [object]$Value
    )

    $parent = Split-Path -Parent $Path
    New-Item -ItemType Directory -Force -Path $parent | Out-Null
    $temporary = "$Path.tmp.$([Guid]::NewGuid().ToString('N'))"
    $Value | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $temporary -Encoding UTF8
    Move-Item -Force -LiteralPath $temporary -Destination $Path
}

function Assert-SkillFolder {
    param(
        [string]$Name,
        [string]$Path
    )

    if ($Name -notmatch '^[a-z0-9-]{1,64}$') {
        throw "Invalid Skill name in release.json: $Name"
    }
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        throw "Skill folder is missing: $Path"
    }
    $skillFile = Join-Path $Path "SKILL.md"
    if (-not (Test-Path -LiteralPath $skillFile -PathType Leaf)) {
        throw "SKILL.md is missing for $Name"
    }
    $content = Get-Content -Raw -LiteralPath $skillFile -Encoding UTF8
    $match = [regex]::Match($content, '(?ms)^\s*---\s*$.*?^\s*name:\s*["'']?([a-z0-9-]+)')
    if (-not $match.Success -or $match.Groups[1].Value -ne $Name) {
        throw "SKILL.md name does not match folder name: $Name"
    }
}

$resolvedCodexHome = Resolve-CodexHome -RequestedHome $CodexHome
$syncRoot = Join-Path $resolvedCodexHome "skill-sync"
$configPath = Join-Path $syncRoot "config.json"
$statePath = Join-Path $syncRoot "state.json"
$skillsRoot = Join-Path $resolvedCodexHome "skills"
$mutexBytes = [Text.Encoding]::UTF8.GetBytes($resolvedCodexHome.ToLowerInvariant())
$mutexHasher = [Security.Cryptography.SHA256]::Create()
$mutexHash = -join ($mutexHasher.ComputeHash($mutexBytes)[0..7] | ForEach-Object { $_.ToString("x2") })
$mutexHasher.Dispose()
$mutex = New-Object Threading.Mutex($false, "Local\SharedSkillSync-$mutexHash")
$hasMutex = $false

try {
    $hasMutex = $mutex.WaitOne(0)
    if (-not $hasMutex) {
        if (-not $HookMode) {
            Write-Host "Another Skill update check is already running."
        }
        return
    }

    $repositoryUrl = $RepositoryUrl
    $branch = $Branch
    if (Test-Path -LiteralPath $configPath -PathType Leaf) {
        $config = Read-JsonFile -Path $configPath
        if (-not [string]::IsNullOrWhiteSpace([string]$config.repository_url)) {
            $repositoryUrl = [string]$config.repository_url
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$config.branch)) {
            $branch = [string]$config.branch
        }
    }
    if ([string]::IsNullOrWhiteSpace($repositoryUrl)) {
        throw "repository_url is empty in $configPath"
    }
    if ([string]::IsNullOrWhiteSpace($branch)) {
        $branch = "main"
    }

    New-Item -ItemType Directory -Force -Path $syncRoot, $skillsRoot | Out-Null
    $legacyRepositoryCache = Join-Path $syncRoot "repository.git"
    if (Test-Path -LiteralPath $legacyRepositoryCache -PathType Container) {
        Remove-Item -Recurse -Force -LiteralPath $legacyRepositoryCache
    }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $github = Resolve-GithubRepository -RepositoryUrl $repositoryUrl
    $owner = [Uri]::EscapeDataString($github.Owner)
    $repository = [Uri]::EscapeDataString($github.Repository)
    $escapedBranch = [Uri]::EscapeDataString($branch)
    $state = $null
    if (Test-Path -LiteralPath $statePath -PathType Leaf) {
        $state = Read-JsonFile -Path $statePath
    }

    $temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) ("codex-skill-sync-" + [Guid]::NewGuid().ToString("N"))
    $archivePath = Join-Path $temporaryRoot "release.zip"
    $extractRoot = Join-Path $temporaryRoot "release"
    New-Item -ItemType Directory -Force -Path $temporaryRoot, $extractRoot | Out-Null

    try {
        $archiveUrl = "https://codeload.github.com/$owner/$repository/zip/refs/heads/$escapedBranch"
        [void](Invoke-PublicDownload -Uri $archiveUrl -OutFile $archivePath)
        Expand-Archive -LiteralPath $archivePath -DestinationPath $extractRoot -Force

        $releaseRoot = $extractRoot
        if (-not (Test-Path -LiteralPath (Join-Path $releaseRoot "release.json") -PathType Leaf)) {
            $releaseRoot = Get-ChildItem -LiteralPath $extractRoot -Directory |
                Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName "release.json") -PathType Leaf } |
                Select-Object -First 1 -ExpandProperty FullName
        }
        if ([string]::IsNullOrWhiteSpace($releaseRoot)) {
            throw "The downloaded GitHub archive does not contain release.json."
        }

        $manifestPath = Join-Path $releaseRoot "release.json"
        $manifest = Read-JsonFile -Path $manifestPath
        if ([int]$manifest.schema_version -ne 1) {
            throw "Unsupported release.json schema version: $($manifest.schema_version)"
        }
        $remoteVersion = [string]$manifest.release_version
        if ([string]::IsNullOrWhiteSpace($remoteVersion)) {
            throw "The downloaded release.json does not contain release_version."
        }
        if (-not $Force -and $null -ne $state -and [string]$state.release_version -eq $remoteVersion) {
            if (-not $HookMode) {
                Write-Host "Skills are already current at release $remoteVersion."
            }
            return
        }

        $skillProperties = @($manifest.skills.PSObject.Properties)
        if ($skillProperties.Count -eq 0) {
            throw "release.json does not list any Skills."
        }

        $releaseRootPrefix = [IO.Path]::GetFullPath($releaseRoot).TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
        $transactionRoot = Join-Path $skillsRoot (".skill-sync-" + [Guid]::NewGuid().ToString("N"))
        $newRoot = Join-Path $transactionRoot "new"
        $oldRoot = Join-Path $transactionRoot "old"
        New-Item -ItemType Directory -Force -Path $newRoot, $oldRoot | Out-Null

        $skillNames = New-Object System.Collections.Generic.List[string]
        foreach ($property in $skillProperties) {
            $name = [string]$property.Name
            $relativePath = [string]$property.Value
            $sourcePath = [IO.Path]::GetFullPath((Join-Path $releaseRoot $relativePath))
            if (-not $sourcePath.StartsWith($releaseRootPrefix, [StringComparison]::OrdinalIgnoreCase)) {
                throw "Skill path escapes the release root: $relativePath"
            }
            Assert-SkillFolder -Name $name -Path $sourcePath

            $stagedPath = Join-Path $newRoot $name
            Copy-Item -Recurse -Force -LiteralPath $sourcePath -Destination $stagedPath
            Get-ChildItem -Recurse -Force -LiteralPath $stagedPath -Directory -Filter "__pycache__" |
                Remove-Item -Recurse -Force
            Get-ChildItem -Recurse -Force -LiteralPath $stagedPath -File -Filter "*.pyc" |
                Remove-Item -Force
            Assert-SkillFolder -Name $name -Path $stagedPath
            $skillNames.Add($name)
        }

        $installedTargets = New-Object System.Collections.Generic.List[string]
        $backups = @{}
        try {
            foreach ($name in $skillNames) {
                $targetPath = Join-Path $skillsRoot $name
                $backupPath = Join-Path $oldRoot $name
                if (Test-Path -LiteralPath $targetPath) {
                    Move-Item -LiteralPath $targetPath -Destination $backupPath
                    $backups[$name] = $backupPath
                }
                Move-Item -LiteralPath (Join-Path $newRoot $name) -Destination $targetPath
                $installedTargets.Add($targetPath)
            }
        } catch {
            foreach ($targetPath in $installedTargets) {
                if (Test-Path -LiteralPath $targetPath) {
                    Remove-Item -Recurse -Force -LiteralPath $targetPath
                }
            }
            foreach ($name in $backups.Keys) {
                $targetPath = Join-Path $skillsRoot $name
                if (Test-Path -LiteralPath $targetPath) {
                    Remove-Item -Recurse -Force -LiteralPath $targetPath
                }
                Move-Item -LiteralPath $backups[$name] -Destination $targetPath
            }
            throw
        }

        $newState = [ordered]@{
            commit = $remoteVersion
            release_version = [string]$manifest.release_version
            updated_at = (Get-Date).ToString("o")
            source = "$($github.Owner)/$($github.Repository)@$branch"
            skills = @($skillNames)
        }
        Write-JsonFileAtomic -Path $statePath -Value $newState
        Remove-Item -Recurse -Force -LiteralPath $transactionRoot

        $namesText = ($skillNames -join ", ")
        if ($HookMode) {
            Write-HookJson -SystemMessage "Skill 自动同步完成：$namesText（$remoteVersion）" -AdditionalContext "The SessionStart hook updated these Skills before this task: $namesText. Before using any of them, read the current SKILL.md from disk again and follow the updated instructions."
        } else {
            Write-Host "Updated Skills: $namesText"
            Write-Host "Release: $($manifest.release_version)"
            Write-Host "Source: $($github.Owner)/$($github.Repository)@$branch"
        }
    } finally {
        if (Test-Path -LiteralPath $temporaryRoot) {
            Remove-Item -Recurse -Force -LiteralPath $temporaryRoot
        }
    }
} catch {
    $message = $_.Exception.Message
    if ($HookMode) {
        Write-HookJson -SystemMessage "Skill 自动同步失败，已保留并继续使用本机旧版本。原因：$message" -AdditionalContext "Skill auto-sync failed at SessionStart. Continue with the installed versions and do not assume an update was applied."
    } else {
        Write-Error $message
        exit 1
    }
} finally {
    if ($hasMutex) {
        $mutex.ReleaseMutex()
    }
    $mutex.Dispose()
}

