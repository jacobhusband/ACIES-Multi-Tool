param(
    [string]$Notes = "Automated release",
    [switch]$Build
)

$ErrorActionPreference = "Stop"

# --- Config ---
$repo = "jacobhusband/ACIES-Multi-Tool"
$assetName = "acies-scheduler-setup.exe"
$versionPath = Join-Path $PSScriptRoot "VERSION"
$repoRootVersionPath = Join-Path (Split-Path $PSScriptRoot -Parent) "VERSION"
$assetPath = Join-Path $PSScriptRoot "dist\\setup\\$assetName"
$repoRootAssetPath = Join-Path (Split-Path $PSScriptRoot -Parent) "dist\\setup\\$assetName"
$repoRootEnvPath = Join-Path (Split-Path $PSScriptRoot -Parent) ".env"

function Get-DotEnvValue {
    param(
        [string]$EnvPath,
        [string]$Name
    )

    if (!(Test-Path $EnvPath)) {
        return $null
    }

    foreach ($line in Get-Content -Path $EnvPath) {
        $trimmedLine = $line.Trim()
        if (-not $trimmedLine -or $trimmedLine.StartsWith("#")) {
            continue
        }

        $match = [regex]::Match($trimmedLine, '^(?<key>[A-Za-z_][A-Za-z0-9_]*)\s*=\s*(?<value>.*)$')
        if (-not $match.Success -or $match.Groups["key"].Value -ne $Name) {
            continue
        }

        $value = $match.Groups["value"].Value.Trim()
        if ($value.Length -ge 2) {
            $firstChar = $value[0]
            $lastChar = $value[$value.Length - 1]
            if (($firstChar -eq '"' -and $lastChar -eq '"') -or ($firstChar -eq "'" -and $lastChar -eq "'")) {
                $value = $value.Substring(1, $value.Length - 2)
            }
        }

        return $value
    }

    return $null
}

# --- Token ---
$token = $env:GITHUB_TOKEN
if (-not $token) {
    $token = Get-DotEnvValue -EnvPath $repoRootEnvPath -Name "GITHUB_TOKEN"
    if ($token) {
        $env:GITHUB_TOKEN = $token
    }
}
if (-not $token) {
    Write-Error "Set GITHUB_TOKEN in the current environment or repo-root .env to a PAT with repo permissions."
    exit 1
}

# --- Version ---
if (!(Test-Path $versionPath)) {
    if (Test-Path $repoRootVersionPath) {
        $versionPath = $repoRootVersionPath
    } else {
        Write-Error "VERSION file not found at $versionPath or $repoRootVersionPath"
        exit 1
    }
}
$version = (Get-Content -Raw $versionPath).Trim()
if (-not $version) {
    Write-Error "VERSION file is empty."
    exit 1
}
$tag = "v$version"

# --- Resolve asset path ---
if (!(Test-Path $assetPath)) {
    if (Test-Path $repoRootAssetPath) {
        $assetPath = $repoRootAssetPath
    }
}

# --- Version check: extract version from built installer and compare to VERSION file ---
function Get-InstallerVersion {
    param([string]$InstallerPath)
    if (!(Test-Path $InstallerPath)) { return $null }
    try {
        $versionInfo = (Get-Item $InstallerPath).VersionInfo
        return $versionInfo.ProductVersion
    } catch {
        return $null
    }
}

$needsBuild = $Build
if (!$needsBuild -and (Test-Path $assetPath)) {
    $installerVersion = Get-InstallerVersion -InstallerPath $assetPath
    if ($installerVersion -and $installerVersion -ne $version) {
        Write-Host "Version mismatch detected:" -ForegroundColor Yellow
        Write-Host "  VERSION file: $version" -ForegroundColor Yellow
        Write-Host "  Installer:    $installerVersion" -ForegroundColor Yellow
        Write-Host "Triggering automatic rebuild..." -ForegroundColor Yellow
        $needsBuild = $true
    } elseif (!$installerVersion) {
        Write-Host "Could not determine installer version. Triggering rebuild to be safe..." -ForegroundColor Yellow
        $needsBuild = $true
    } else {
        Write-Host "Version check passed: $version" -ForegroundColor Green
    }
} elseif (!(Test-Path $assetPath)) {
    Write-Host "Installer not found. Triggering build..." -ForegroundColor Yellow
    $needsBuild = $true
}

# --- Build if needed ---
if ($needsBuild) {
    & "$PSScriptRoot/build.ps1"
    if ($LASTEXITCODE -ne 0) { exit 1 }
    # Re-resolve asset path after build
    if (!(Test-Path $assetPath)) {
        if (Test-Path $repoRootAssetPath) {
            $assetPath = $repoRootAssetPath
        } else {
            Write-Error "Installer not found after build at $assetPath or $repoRootAssetPath"
            exit 1
        }
    }
}

$headers = @{
    "Authorization" = "token $token"
    "Accept"        = "application/vnd.github+json"
    "User-Agent"    = "publish-release-script"
}

# --- Create release if missing ---
$release = $null
try {
    $release = Invoke-RestMethod -Method Get -Headers $headers -Uri "https://api.github.com/repos/$repo/releases/tags/$tag"
} catch {
    $release = $null
}

if (-not $release) {
    $body = @{
        tag_name   = $tag
        name       = $tag
        body       = $Notes
        draft      = $false
        prerelease = $false
    } | ConvertTo-Json
    $release = Invoke-RestMethod -Method Post -Headers $headers -Uri "https://api.github.com/repos/$repo/releases" -Body $body
}

if (-not $release.upload_url) {
    Write-Error "Release upload_url missing. Check token permissions (repo: write) and try again."
    exit 1
}

# Normalize upload URL from release response
$uploadUrl = [string]$release.upload_url
$uploadUrl = $uploadUrl.Replace("{?name,label}", "").Trim()

$encodedName = [uri]::EscapeDataString($assetName)
$uploadUri = [string]::Format("{0}?name={1}", $uploadUrl, $encodedName)

# Basic validation/diagnostics for URI issues
if ([string]::IsNullOrWhiteSpace($uploadUrl)) {
    Write-Error "upload_url is empty. Release response: $($release | ConvertTo-Json -Depth 4)"
    exit 1
}

try { [void][uri]$uploadUri } catch {
    Write-Error "uploadUri is invalid: $uploadUri`nRelease upload_url: $uploadUrl`nError: $_"
    exit 1
}

# --- Upload asset (clobber if exists) ---
try {
    $assets = Invoke-RestMethod -Method Get -Headers $headers -Uri "https://api.github.com/repos/$repo/releases/$($release.id)/assets"
    foreach ($a in $assets) {
        if ($a.name -eq $assetName) {
            Invoke-RestMethod -Method Delete -Headers $headers -Uri "https://api.github.com/repos/$repo/releases/assets/$($a.id)" | Out-Null
        }
    }
} catch { }

$assetHeaders = $headers.Clone()
$assetHeaders["Content-Type"] = "application/octet-stream"
try {
    Write-Host "Uploading asset to: $uploadUri" -ForegroundColor Cyan
    Invoke-RestMethod -Method Post -Headers $assetHeaders -Uri $uploadUri -InFile $assetPath -TimeoutSec 600 | Out-Null
} catch {
    Write-Error "Asset upload failed. uploadUri: $uploadUri`nRelease upload_url: $uploadUrl`nError: $_"
    exit 1
}

Write-Host "Release $tag published with asset $assetName" -ForegroundColor Green
