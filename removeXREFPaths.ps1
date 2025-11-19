# removeXrefPaths.ps1
# Select DWGs -> accoreconsole loads your .NET DLL via NETLOAD -> runs STRIPREFPATHS
# All feedback is sent to the console for the main application to display.

Write-Host "PROGRESS: Initializing script..."

# --- AUTOCAD VERSION DETECTION (2025-2020) ---
$acadCore = $null
$years = 2025, 2024, 2023, 2022, 2021, 2020

foreach ($year in $years) {
    $possiblePath = "C:\Program Files\Autodesk\AutoCAD $year\accoreconsole.exe"
    if (Test-Path -Path $possiblePath) {
        $acadCore = $possiblePath
        Write-Host "PROGRESS: Found AutoCAD $year Core Console."
        break # Stop searching once the latest version is found
    }
}

$scriptRoot = $PSScriptRoot
$dll = Join-Path $scriptRoot "StripRefPaths.dll"

# Relaunch in STA for file picker
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
  $ps = (Get-Process -Id $PID).Path
  Start-Process -FilePath $ps -ArgumentList @("-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-File", "`"$PSCommandPath`"") -Wait
  exit
}

# Validation
if ([string]::IsNullOrEmpty($acadCore) -or -not (Test-Path $acadCore)) { 
    Write-Host "PROGRESS: ERROR: AutoCAD Core Console not found for versions 2020-2025."
    Write-Host "Please ensure AutoCAD is installed in the default 'C:\Program Files\Autodesk' directory."
    exit 1 
}

if (-not (Test-Path $dll)) { 
    Write-Host "PROGRESS: ERROR: Required component 'StripRefPaths.dll' not found at $dll"
    exit 1 
}

Write-Host "PROGRESS: Staging cleanup commands..."
# Stage AutoLISP CLEANUP command so each drawing is tidied after ref paths are stripped
$lisp = Join-Path $env:TEMP "cleanup_tmp.lsp"
$lispContent = @'
(defun c:CLEANUP (/ oldCMDECHO oldTab)
  ;; Remember current command echo setting, then turn it off:
  (setq oldCMDECHO (getvar "CMDECHO")
        oldTab     (getvar "CTAB"))
  (setvar "CMDECHO" 0)

  ;; Ensure commands run from Model space
  (if (/= oldTab "Model")
    (command "_.MODEL")
  )

  ;; Use SETBYLAYER on everything in the current space
  ;;   "Y" => Change ByBlock to ByLayer?
  ;;   "Y" => Include blocks?
  (command
    "_.-SETBYLAYER" "All" "" "Y" "Y")

  ;; Purge the entire drawing; "All", "*" (everything), "No" to confirm each
  (command "_.-PURGE" "All" "*" "N")

  ;; Audit the drawing; "Yes" to fix errors
  (command "_.AUDIT" "Y")

  ;; Restore layout/model state if it changed
  (if (/= oldTab (getvar "CTAB"))
    (if (= oldTab "Model")
      (command "_.MODEL")
      (command "_.LAYOUT" "Set" oldTab)
    )
  )

  ;; Restore command echo setting
  (setvar "CMDECHO" oldCMDECHO)
  (princ)
)
'@
Set-Content -Encoding ASCII -Path $lisp -Value $lispContent

# Make .scr that NETLOADs the DLL, strips refs, and then runs CLEANUP
$script = Join-Path $env:TEMP "run_STRIP.scr"
$lispPathForScript = ($lisp -replace '\\', '/')
$scriptContent = @"
CMDECHO 1
FILEDIA 0
SECURELOAD 0
NETLOAD
"$dll"
STRIPREFPATHS
(load "$lispPathForScript")
CLEANUP
QSAVE
QUIT
"@
Set-Content -Encoding ASCII -Path $script -Value $scriptContent

# Pick DWGs
Write-Host "PROGRESS: Waiting for file selection..."
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
$dlg = New-Object System.Windows.Forms.OpenFileDialog
$dlg.Title = "Select DWG file(s) to strip saved XREF/underlay/image paths"
$dlg.Filter = "DWG files (*.dwg)|*.dwg|All files (*.*)|*.*"
$dlg.Multiselect = $true
$dlg.InitialDirectory = [Environment]::GetFolderPath("Desktop")
if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK -or -not $dlg.FileNames) {
  Write-Host "PROGRESS: ERROR: No files selected."; exit 1
}
$files = $dlg.FileNames

$failed = @()
$i = 0
foreach ($dwg in $files) {
  $i++
  $name = [IO.Path]::GetFileName($dwg)
  Write-Host "PROGRESS: Processing $i of $($files.Count): $name"

  # IMPORTANT: no /ld here; we NETLOAD via the .scr
  # The 2>&1 redirects stderr to stdout, so the Python wrapper can catch errors.
  & $acadCore /i "$dwg" /s "$script" 2>&1
  $code = $LASTEXITCODE
  if ($code -ne 0) { $failed += $dwg }
}

Write-Progress -Activity "Processing DWGs" -Completed

if ($failed.Count) {
  Write-Host "PROGRESS: ERROR: $($failed.Count) of $($files.Count) file(s) failed to process."
} else {
  Write-Host "PROGRESS: Successfully processed $($files.Count) drawing(s)."
}