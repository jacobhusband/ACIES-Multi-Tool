param(
    [string]$Notes = "Automated release",
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"

$versionPath = Join-Path $PSScriptRoot "VERSION"
$repoRootVersionPath = Join-Path (Split-Path $PSScriptRoot -Parent) "VERSION"
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

# Optional: ensure working tree is clean
$gitStatus = git status --porcelain
if ($gitStatus) {
    Write-Warning "Git working tree is not clean. Commit/stash changes before releasing."
    Write-Warning $gitStatus
    exit 1
}

if (-not $SkipBuild) {
    Write-Host "Building installer for version $version..." -ForegroundColor Cyan
    & "$PSScriptRoot/build.ps1"
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Build failed."
        exit 1
    }
}

$assetPath = Join-Path $PSScriptRoot "dist/setup/acies-scheduler-setup.exe"
$repoRootAssetPath = Join-Path (Split-Path $PSScriptRoot -Parent) "dist/setup/acies-scheduler-setup.exe"
if (!(Test-Path $assetPath)) {
    if (Test-Path $repoRootAssetPath) {
        $assetPath = $repoRootAssetPath
    } else {
        Write-Error "Installer not found at $assetPath or $repoRootAssetPath"
        exit 1
    }
}

# Create git tag if it doesn't exist
$existingTag = git tag --list $tag
if (-not $existingTag) {
    git tag $tag
    git push origin $tag
} else {
    Write-Host "Tag $tag already exists; skipping tag creation." -ForegroundColor Yellow
}

# Create or update GitHub release using gh CLI
if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    Write-Error "'gh' CLI is not installed. Install from https://github.com/cli/cli."
    exit 1
}

# Try to create; if exists, upload asset
$releaseExists = $false
try {
    gh release view $tag | Out-Null
    $releaseExists = $true
} catch {
    $releaseExists = $false
}

if (-not $releaseExists) {
    gh release create $tag $assetPath --notes "$Notes" --title $tag
} else {
    gh release upload $tag $assetPath --clobber
    gh release edit $tag --notes "$Notes" --title $tag
}

Write-Host "Release $tag published with asset $assetPath" -ForegroundColor Green
