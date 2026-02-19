Write-Host "###################################" -ForegroundColor Cyan
Write-Host "#      Building ACIES Scheduler     #" -ForegroundColor Cyan
Write-Host "###################################" -ForegroundColor Cyan

# Project root is one level up from build-config
$projectRoot = Split-Path $PSScriptRoot -Parent

# 1. Build Wire Sizer frontend (if present)
$wireSizerDir = Join-Path $projectRoot "WireSizerApplication"
$wireSizerDist = Join-Path $wireSizerDir "dist"
if (Test-Path $wireSizerDir) {
    Write-Host "`n[1/3] Building Wire Sizer frontend..." -ForegroundColor Yellow
    $npmCmd = Get-Command npm -ErrorAction SilentlyContinue
    if (-not $npmCmd) {
        Write-Warning "npm not found. Skipping Wire Sizer build."
    } else {
        Push-Location $wireSizerDir
        if (Test-Path "package.json") {
            if (-not (Test-Path "node_modules")) {
                npm install
            }
            npm run build
        } else {
            Write-Warning "WireSizerApplication package.json not found. Skipping build."
        }
        Pop-Location
    }
} else {
    Write-Warning "WireSizerApplication folder not found. Skipping build."
}

# 2. Run PyInstaller
Write-Host "`n[2/3] Bundling application with PyInstaller..." -ForegroundColor Yellow

# Clean previous installer output to avoid file locks/resource errors
$setupOutput = Join-Path $projectRoot "dist\\setup"
if (Test-Path $setupOutput) {
    try {
        Remove-Item -Recurse -Force $setupOutput
    } catch {
        Write-Warning "Could not clean dist\\setup: $_"
    }
}

# Read version from VERSION file
$versionPath = Join-Path $projectRoot "VERSION"
if (!(Test-Path $versionPath)) {
    Write-Error "VERSION file not found at $versionPath"
    exit 1
}
$appVersion = (Get-Content -Raw $versionPath).Trim()
if (-not $appVersion) {
    Write-Error "VERSION file is empty."
    exit 1
}

# Change to project root for PyInstaller
Push-Location $projectRoot

# --noconfirm: Overwrites output directory without asking
# --clean: Cleans PyInstaller cache
$pyInstallerArgs = @(
    "main.py", "--noconfirm", "--clean", "--noconsole",
    "--name", "ACIES Scheduler",
    "--icon=assets\acies.ico",
    "--add-data", "VERSION;.",
    "--add-data", "index.html;.",
    "--add-data", "styles.css;.",
    "--add-data", "script.js;.",
    "--add-data", ".env;.",
    "--add-data", "assets\acies.png;.",
    "--add-data", "scripts\merge_pdfs.py;scripts",
    "--add-data", "scripts\detect_pdf_size.py;scripts",
    "--add-data", "scripts\PlotDWGs.ps1;scripts",
    "--add-data", "scripts\FreezeLayersDWGs.ps1;scripts",
    "--add-data", "scripts\ThawLayersDWGs.ps1;scripts",
    "--add-data", "scripts\removeXREFPaths.ps1;scripts",
    "--add-data", "scripts\StripRefPaths.dll;scripts",
    "--add-data", "templates;templates"
)

if (Test-Path $wireSizerDist) {
    $pyInstallerArgs += @("--add-data", "WireSizerApplication\\dist;WireSizerApplication\\dist")
} else {
    Write-Warning "WireSizerApplication\\dist not found. Wire Sizer tool will be unavailable."
}

# Run PyInstaller and check for errors
& pyinstaller $pyInstallerArgs
if ($LASTEXITCODE -ne 0) {
    Pop-Location
    Write-Error "PyInstaller failed!"
    exit 1
}

Pop-Location

# 3. Run Inno Setup
Write-Host "`n[3/3] Creating installer with Inno Setup..." -ForegroundColor Yellow

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

# Run Inno Setup (setup.iss is in the same build-config folder)
& $isccPath "/DMyAppVersion=$appVersion" (Join-Path $PSScriptRoot "setup.iss")
if ($LASTEXITCODE -ne 0) {
    Write-Error "Inno Setup compilation failed!"
    exit 1
}

Write-Host "`n###################################" -ForegroundColor Green
Write-Host "#      Build process complete!      #" -ForegroundColor Green
Write-Host "###################################" -ForegroundColor Green
Write-Host "Installer created at: dist\setup\acies-scheduler-setup.exe"
