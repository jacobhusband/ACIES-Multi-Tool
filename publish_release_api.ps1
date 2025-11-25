param(
    [string]$Notes = "Automated release",
    [switch]$Build
)

$ErrorActionPreference = "Stop"

# --- Config ---
$repo = "jacobhusband/ACIES-Multi-Tool"
$assetName = "acies-scheduler-setup.exe"
$versionPath = Join-Path $PSScriptRoot "VERSION"
$assetPath = Join-Path $PSScriptRoot "dist\\setup\\$assetName"

# --- Token ---
$token = $env:GITHUB_TOKEN
if (-not $token) {
    Write-Error "Set environment variable GITHUB_TOKEN to a PAT with repo permissions."
    exit 1
}

# --- Version ---
if (!(Test-Path $versionPath)) {
    Write-Error "VERSION file not found at $versionPath"
    exit 1
}
$version = (Get-Content -Raw $versionPath).Trim()
if (-not $version) {
    Write-Error "VERSION file is empty."
    exit 1
}
$tag = "v$version"

# --- Optional build ---
if ($Build) {
    & "$PSScriptRoot/build.ps1"
    if ($LASTEXITCODE -ne 0) { exit 1 }
}

if (!(Test-Path $assetPath)) {
    Write-Error "Installer not found at $assetPath. Build first or pass -Build."
    exit 1
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
