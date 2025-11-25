Write-Host "###################################" -ForegroundColor Cyan
Write-Host "#      Building ACIES Scheduler     #" -ForegroundColor Cyan
Write-Host "###################################" -ForegroundColor Cyan

# 1. Run PyInstaller
Write-Host "`n[1/2] Bundling application with PyInstaller..." -ForegroundColor Yellow

# Clean previous installer output to avoid file locks/resource errors
$setupOutput = Join-Path $PSScriptRoot "dist\\setup"
if (Test-Path $setupOutput) {
    try {
        Remove-Item -Recurse -Force $setupOutput
    } catch {
        Write-Warning "Could not clean dist\\setup: $_"
    }
}

# Read version from VERSION file
$versionPath = Join-Path $PSScriptRoot "VERSION"
if (!(Test-Path $versionPath)) {
    Write-Error "VERSION file not found at $versionPath"
    exit 1
}
$appVersion = (Get-Content -Raw $versionPath).Trim()
if (-not $appVersion) {
    Write-Error "VERSION file is empty."
    exit 1
}

# --noconfirm: Overwrites output directory without asking
# --clean: Cleans PyInstaller cache
$pyInstallerArgs = @(
    "main.py", "--noconfirm", "--clean", "--noconsole", 
    "--name", "ACIES Scheduler", 
    "--icon=acies.ico",
    "--add-data", "VERSION;.",
    "--add-data", "index.html;.",
    "--add-data", "styles.css;.",
    "--add-data", "script.js;.",
    "--add-data", ".env;.",
    "--add-data", "acies.png;.",
    "--add-data", "merge_pdfs.py;.",
    "--add-data", "PlotDWGs.ps1;.",
    "--add-data", "CleanDWGs.ps1;.",
    "--add-data", "removeXREFPaths.ps1;.",
    "--add-data", "StripRefPaths.dll;."
)

# Run PyInstaller and check for errors
& pyinstaller $pyInstallerArgs
if ($LASTEXITCODE -ne 0) {
    Write-Error "PyInstaller failed!"
    exit 1
}

# 2. Run Inno Setup
Write-Host "`n[2/2] Creating installer with Inno Setup..." -ForegroundColor Yellow

# Define potential paths for ISCC.exe
# Note: We look for ISCC.exe (Command Line Compiler), not Compil32.exe (GUI)
$possiblePaths = @(
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",           # Explicitly requested location
    "$env:ProgramFiles(x86)\Inno Setup 6\iscc.exe",           # Standard x64 install (Env Var)
    "$env:ProgramFiles\Inno Setup 6\iscc.exe",                # Standard x86 install
    "$env:LOCALAPPDATA\Programs\Inno Setup 6\iscc.exe"        # User-level install
)

$isccPath = $null
foreach ($path in $possiblePaths) {
    if (Test-Path $path) {
        $isccPath = $path
        break
    }
}

if (-not $isccPath) {
    Write-Error "Inno Setup Compiler (iscc.exe) not found."
    Write-Host "Checked paths:"
    $possiblePaths | ForEach-Object { Write-Host " - $_" }
    exit 1
}

Write-Host "Found Inno Setup at: $isccPath" -ForegroundColor Gray

# Run Inno Setup
& $isccPath "/DMyAppVersion=$appVersion" "setup.iss"
if ($LASTEXITCODE -ne 0) {
    Write-Error "Inno Setup compilation failed!"
    exit 1
}

Write-Host "`n###################################" -ForegroundColor Green
Write-Host "#      Build process complete!      #" -ForegroundColor Green
Write-Host "###################################" -ForegroundColor Green
Write-Host "Installer created at: dist\setup\acies-scheduler-setup.exe"
