[CmdletBinding()]
param(
    [string]$RepositoryUrl = "https://github.com/13349811148/skill.git",
    [string]$Branch = "main",
    [string]$CodexHome,
    [string]$WorkBuddyHome,
    [string]$CodeBuddyHome
)

$ErrorActionPreference = "Stop"

function Add-InstallTarget {
    param(
        [System.Collections.Generic.List[object]]$Targets,
        [string]$Name,
        [string]$AgentHome,
        [string]$SettingsFile
    )

    if ([string]::IsNullOrWhiteSpace($AgentHome)) {
        return
    }
    $fullHome = [IO.Path]::GetFullPath($AgentHome)
    if (@($Targets | Where-Object { $_.Home -eq $fullHome }).Count -eq 0) {
        $Targets.Add([pscustomobject]@{
            Name = $Name
            Home = $fullHome
            SettingsPath = Join-Path $fullHome $SettingsFile
        })
    }
}

function Merge-SessionStartHook {
    param(
        [string]$SettingsPath,
        [string]$AgentHome,
        [string]$InstalledUpdater
    )

    if (Test-Path -LiteralPath $SettingsPath -PathType Leaf) {
        $settingsDocument = Get-Content -Raw -LiteralPath $SettingsPath -Encoding UTF8 | ConvertFrom-Json
        $backupPath = "$SettingsPath.backup.$(Get-Date -Format 'yyyyMMddHHmmssfff')"
        Copy-Item -LiteralPath $SettingsPath -Destination $backupPath
    } else {
        $settingsDocument = [pscustomobject]@{}
    }

    if ($null -eq $settingsDocument.PSObject.Properties["hooks"]) {
        $settingsDocument | Add-Member -NotePropertyName "hooks" -NotePropertyValue ([pscustomobject]@{})
    }
    if ($null -eq $settingsDocument.hooks.PSObject.Properties["SessionStart"]) {
        $settingsDocument.hooks | Add-Member -NotePropertyName "SessionStart" -NotePropertyValue @()
    }

    $existingEntries = @($settingsDocument.hooks.SessionStart)
    $filteredEntries = @($existingEntries | Where-Object {
        $commands = @($_.hooks | ForEach-Object { "$($_.command) $($_.commandWindows)" }) -join " "
        $commands -notlike "*skill-sync*update-skills.ps1*"
    })

    $hookCommand = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$InstalledUpdater`" -CodexHome `"$AgentHome`" -HookMode"
    $syncEntry = [pscustomobject]@{
        matcher = "startup|resume"
        hooks = @(
            [pscustomobject]@{
                type = "command"
                command = $hookCommand
                commandWindows = $hookCommand
                timeout = 90
                statusMessage = "Checking shared Skill updates"
            }
        )
    }
    $settingsDocument.hooks.SessionStart = @($filteredEntries + $syncEntry)

    $settingsParent = Split-Path -Parent $SettingsPath
    New-Item -ItemType Directory -Force -Path $settingsParent | Out-Null
    $settingsTemporary = "$SettingsPath.tmp.$([Guid]::NewGuid().ToString('N'))"
    $settingsDocument | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $settingsTemporary -Encoding UTF8
    Move-Item -Force -LiteralPath $settingsTemporary -Destination $SettingsPath
}

function Install-AgentTarget {
    param([pscustomobject]$Target)

    $syncRoot = Join-Path $Target.Home "skill-sync"
    $installedUpdater = Join-Path $syncRoot "update-skills.ps1"
    $configPath = Join-Path $syncRoot "config.json"
    New-Item -ItemType Directory -Force -Path $syncRoot | Out-Null
    Copy-Item -Force -LiteralPath $script:sourceUpdater -Destination $installedUpdater

    $config = [ordered]@{
        repository_url = $script:RepositoryUrl
        branch = $script:Branch
        product = $Target.Name
    }
    $config | ConvertTo-Json | Set-Content -LiteralPath $configPath -Encoding UTF8

    Write-Host "正在为 $($Target.Name) 从 GitHub 安装当前版本..."
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $installedUpdater -CodexHome $Target.Home -Force
    if ($LASTEXITCODE -ne 0) {
        throw "$($Target.Name) initial Skill synchronization failed. Its startup hook was not installed."
    }

    Merge-SessionStartHook -SettingsPath $Target.SettingsPath -AgentHome $Target.Home -InstalledUpdater $installedUpdater
    return [pscustomobject]@{
        Name = $Target.Name
        Home = $Target.Home
        SkillsPath = Join-Path $Target.Home "skills"
        SettingsPath = $Target.SettingsPath
    }
}

if ([string]::IsNullOrWhiteSpace($RepositoryUrl)) {
    throw "GitHub repository URL is required."
}

$sourceUpdater = Join-Path $PSScriptRoot "sync\update-skills.ps1"
if (-not (Test-Path -LiteralPath $sourceUpdater -PathType Leaf)) {
    throw "Updater is missing: $sourceUpdater"
}

$targets = New-Object 'System.Collections.Generic.List[object]'
$hasExplicitHome = $PSBoundParameters.ContainsKey("CodexHome") -or
    $PSBoundParameters.ContainsKey("WorkBuddyHome") -or
    $PSBoundParameters.ContainsKey("CodeBuddyHome")

if ($hasExplicitHome) {
    Add-InstallTarget -Targets $targets -Name "Codex" -AgentHome $CodexHome -SettingsFile "hooks.json"
    Add-InstallTarget -Targets $targets -Name "WorkBuddy" -AgentHome $WorkBuddyHome -SettingsFile "settings.json"
    Add-InstallTarget -Targets $targets -Name "CodeBuddy" -AgentHome $CodeBuddyHome -SettingsFile "settings.json"
} else {
    $defaultCodexHome = if (-not [string]::IsNullOrWhiteSpace($env:CODEX_HOME)) {
        $env:CODEX_HOME
    } else {
        Join-Path $env:USERPROFILE ".codex"
    }
    $defaultWorkBuddyHome = Join-Path $env:USERPROFILE ".workbuddy"
    $defaultCodeBuddyHome = Join-Path $env:USERPROFILE ".codebuddy"

    if (Test-Path -LiteralPath $defaultCodexHome -PathType Container) {
        Add-InstallTarget -Targets $targets -Name "Codex" -AgentHome $defaultCodexHome -SettingsFile "hooks.json"
    }
    if (Test-Path -LiteralPath $defaultWorkBuddyHome -PathType Container) {
        Add-InstallTarget -Targets $targets -Name "WorkBuddy" -AgentHome $defaultWorkBuddyHome -SettingsFile "settings.json"
    }
    if (Test-Path -LiteralPath $defaultCodeBuddyHome -PathType Container) {
        Add-InstallTarget -Targets $targets -Name "CodeBuddy" -AgentHome $defaultCodeBuddyHome -SettingsFile "settings.json"
    }
}

if ($targets.Count -eq 0) {
    throw "No supported AI client was detected. Open Codex, WorkBuddy, or CodeBuddy once, then run this installer again."
}

$installedTargets = New-Object 'System.Collections.Generic.List[object]'
foreach ($target in $targets) {
    $installedTargets.Add((Install-AgentTarget -Target $target))
}

Write-Host ""
Write-Host "安装完成。"
Write-Host "GitHub: $RepositoryUrl"
Write-Host "分支: $Branch"
foreach ($installed in $installedTargets) {
    Write-Host "$($installed.Name) Skills: $($installed.SkillsPath)"
    Write-Host "$($installed.Name) 启动配置: $($installed.SettingsPath)"
}
Write-Host "请完全退出并重新打开对应软件。首次出现钩子审核或信任提示时，请确认信任此同步钩子。"


