$ErrorActionPreference = "Stop"

function Invoke-CheckedCommand {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Description,
        [Parameter(Mandatory = $true)]
        [scriptblock]$Command
    )

    $global:LASTEXITCODE = 0
    & $Command
    if ($LASTEXITCODE -ne 0) {
        throw "$Description failed with exit code $LASTEXITCODE."
    }
}

function Assert-PathExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if (-not (Test-Path $Path)) {
        throw $Message
    }
}

try {
    Write-Host "###################################" -ForegroundColor Cyan
    Write-Host "#      Building ACIES Scheduler     #" -ForegroundColor Cyan
    Write-Host "###################################" -ForegroundColor Cyan

    # Pin the build to the repo-local virtualenv so editor-integrated shells do not change behavior.
    $projectRoot = Split-Path $PSScriptRoot -Parent
    $venvPython = Join-Path $projectRoot ".venv\Scripts\python.exe"
    Assert-PathExists -Path $venvPython -Message "Repo-local Python interpreter not found at $venvPython. Create or refresh .venv before building."
    Write-Host "Using Python: $venvPython" -ForegroundColor Gray
    Invoke-CheckedCommand -Description "PyInstaller availability check" -Command { & $venvPython -m PyInstaller --version | Out-Null }

    $wireSizerDir = Join-Path $projectRoot "WireSizerApplication"
    $wireSizerPackageJson = Join-Path $wireSizerDir "package.json"
    $wireSizerDist = Join-Path $wireSizerDir "dist"
    $wireSizerIndex = Join-Path $wireSizerDist "index.html"
    Assert-PathExists -Path $wireSizerDir -Message "WireSizerApplication folder not found at $wireSizerDir."
    Assert-PathExists -Path $wireSizerPackageJson -Message "Wire Sizer package.json not found at $wireSizerPackageJson."

    $npmCmd = Get-Command npm -ErrorAction SilentlyContinue
    if (-not $npmCmd) {
        throw "npm not found on PATH. Install Node.js/npm before building."
    }
    $npmExecutable = $npmCmd.Source

    Write-Host "`n[1/3] Building Wire Sizer frontend..." -ForegroundColor Yellow
    Write-Host "Using npm: $npmExecutable" -ForegroundColor Gray
    Push-Location $wireSizerDir
    try {
        if (-not (Test-Path "node_modules")) {
            Invoke-CheckedCommand -Description "npm install" -Command { & $npmExecutable install }
        }
        Invoke-CheckedCommand -Description "npm run build" -Command { & $npmExecutable run build }
    } finally {
        Pop-Location
    }
    Assert-PathExists -Path $wireSizerDist -Message "Wire Sizer build completed without producing $wireSizerDist."
    Assert-PathExists -Path $wireSizerIndex -Message "Wire Sizer build output is missing $wireSizerIndex."

    Write-Host "`n[2/3] Bundling application with PyInstaller..." -ForegroundColor Yellow

    $bundleOutput = Join-Path $projectRoot "dist\ACIES Scheduler"
    $bundleInternal = Join-Path $bundleOutput "_internal"
    $setupOutput = Join-Path $projectRoot "dist\setup"
    foreach ($path in @($bundleOutput, $setupOutput)) {
        if (Test-Path $path) {
            Write-Host "Cleaning $path" -ForegroundColor Gray
            Remove-Item -Recurse -Force $path
        }
    }

    $versionPath = Join-Path $projectRoot "VERSION"
    Assert-PathExists -Path $versionPath -Message "VERSION file not found at $versionPath."
    $appVersion = (Get-Content -Raw $versionPath).Trim()
    if (-not $appVersion) {
        throw "VERSION file is empty."
    }

    Push-Location $projectRoot
    try {
        $pyInstallerSpecPath = "build\pyinstaller"
        $pyInstallerArgs = @(
            "main.py", "--noconfirm", "--clean", "--noconsole",
            "--name", "ACIES Scheduler",
            "--specpath", $pyInstallerSpecPath,
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
            "--add-data", "templates;templates",
            "--add-data", "WireSizerApplication\dist;WireSizerApplication\dist"
        )

        Invoke-CheckedCommand -Description "PyInstaller build" -Command { & $venvPython -m PyInstaller $pyInstallerArgs }
    } finally {
        Pop-Location
    }

    $expectedBundleFiles = @(
        (Join-Path $bundleOutput "ACIES Scheduler.exe"),
        (Join-Path $bundleInternal "index.html"),
        (Join-Path $bundleInternal "styles.css"),
        (Join-Path $bundleInternal "script.js"),
        (Join-Path $bundleInternal "WireSizerApplication\dist\index.html")
    )
    foreach ($expectedPath in $expectedBundleFiles) {
        Assert-PathExists -Path $expectedPath -Message "PyInstaller bundle is missing expected output: $expectedPath"
    }

    Write-Host "`n[3/3] Creating installer with Inno Setup..." -ForegroundColor Yellow

    $possiblePaths = @(
        "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
        "$env:ProgramFiles(x86)\Inno Setup 6\iscc.exe",
        "$env:ProgramFiles\Inno Setup 6\iscc.exe",
        "$env:LOCALAPPDATA\Programs\Inno Setup 6\iscc.exe"
    )

    $isccPath = $null
    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            $isccPath = $path
            break
        }
    }

    if (-not $isccPath) {
        throw "Inno Setup Compiler (iscc.exe) not found. Checked paths: $($possiblePaths -join ', ')"
    }

    Write-Host "Found Inno Setup at: $isccPath" -ForegroundColor Gray
    Invoke-CheckedCommand -Description "Inno Setup compilation" -Command {
        & $isccPath "/DMyAppVersion=$appVersion" (Join-Path $PSScriptRoot "setup.iss")
    }

    $installerPath = Join-Path $setupOutput "acies-scheduler-setup.exe"
    Assert-PathExists -Path $installerPath -Message "Inno Setup finished without producing $installerPath."

    Write-Host "`n###################################" -ForegroundColor Green
    Write-Host "#      Build process complete!      #" -ForegroundColor Green
    Write-Host "###################################" -ForegroundColor Green
    Write-Host "Installer created at: dist\setup\acies-scheduler-setup.exe"
} catch {
    Write-Error $_
    exit 1
}
