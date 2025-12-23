param(
  [string]$AcadCore
)

# ---------------- CONFIGURATION ----------------
$ProcessTimeoutSeconds = 120
$ToolDir = Join-Path $env:LOCALAPPDATA "AcadHeadlessTools"

Write-Host "PROGRESS: Initializing script..."

# ---------------- 1) FIND ACCORECONSOLE ----------------
if ($AcadCore -and (Test-Path -Path $AcadCore)) {
  $acadCore = $AcadCore
}
else {
  $acadCore = $null
  $years = 2026..2018
  foreach ($year in $years) {
    $possiblePath = "C:\Program Files\Autodesk\AutoCAD $year\accoreconsole.exe"
    if (Test-Path -Path $possiblePath) {
      $acadCore = $possiblePath
      Write-Host "PROGRESS: Found AutoCAD $year Core Console."
      break
    }
  }
}

if (-not $acadCore) {
  Write-Error "AutoCAD Core Console not found. Provide -AcadCore or install AutoCAD."
  exit 1
}

# ---------------- 2) SELECT FILES (STA wrapper) ----------------
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
  $ps = (Get-Process -Id $PID).Path
  $argsList = @("-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-File", "`"$PSCommandPath`"")
  if ($AcadCore) { $argsList += @("-AcadCore", "`"$AcadCore`"") }

  Start-Process -FilePath $ps -ArgumentList $argsList -Wait
  exit
}

Add-Type -AssemblyName System.Windows.Forms
$dlg = New-Object System.Windows.Forms.OpenFileDialog
$dlg.Title = "Select DWG file(s)"
$dlg.Filter = "DWG files (*.dwg)|*.dwg"
$dlg.Multiselect = $true
if ($dlg.ShowDialog() -ne 'OK') { exit }
$files = $dlg.FileNames

# ---------------- 3) PREP TOOL FOLDER ----------------
if (-not (Test-Path $ToolDir)) { New-Item -ItemType Directory -Path $ToolDir | Out-Null }

function Stop-ProcessTree {
  param([int]$Pid)
  try { & taskkill /PID $Pid /T /F | Out-Null } catch {
    try { Stop-Process -Id $Pid -Force -ErrorAction SilentlyContinue } catch {}
  }
}

function Invoke-AcadCore {
  param(
    [Parameter(Mandatory = $true)][string]$DwgPath,
    [Parameter(Mandatory = $true)][string]$ScriptPath,
    [Parameter(Mandatory = $true)][string]$OutLog,
    [Parameter(Mandatory = $true)][string]$ErrLog,
    [int]$TimeoutSeconds = 120
  )

  $p = Start-Process -FilePath $acadCore `
    -ArgumentList "/i `"$DwgPath`" /s `"$ScriptPath`"" `
    -PassThru -NoNewWindow `
    -RedirectStandardOutput $OutLog `
    -RedirectStandardError  $ErrLog

  if (-not $p.WaitForExit($TimeoutSeconds * 1000)) {
    Stop-ProcessTree -Pid $p.Id
    return @{ TimedOut = $true; ExitCode = $null; Pid = $p.Id }
  }

  return @{ TimedOut = $false; ExitCode = $p.ExitCode; Pid = $p.Id }
}

# ---------------- 4) EXTRACTION PHASE ----------------
$layerDumpFile = Join-Path $ToolDir "layers_dump.txt"
if (Test-Path $layerDumpFile) { Remove-Item $layerDumpFile -Force }

$lispReportPath = ($layerDumpFile -replace '\\', '/')

$extractLsp = Join-Path $ToolDir "extract.lsp"
$extractLspForLisp = ($extractLsp -replace '\\', '/')

# No (quit) here
$extractLspContent = @"
(defun c:ExtractLayers (/ f lay)
  (setq f (open "$lispReportPath" "a"))
  (if f
    (progn
      (write-line (strcat "###DWG:" (getvar "DWGNAME")) f)
      (setq lay (tblnext "LAYER" T))
      (while lay
        (write-line (cdr (assoc 2 lay)) f)
        (setq lay (tblnext "LAYER"))
      )
      (close f)
    )
  )
  (princ)
)
(princ)
"@
Set-Content -Path $extractLsp -Value $extractLspContent -Encoding ASCII

$extractScr = Join-Path $ToolDir "extract.scr"
$extractScrContent = @"
FILEDIA
0
CMDDIA
0
PROXYNOTICE
0
SECURELOAD
0
(load "$extractLspForLisp")
(c:ExtractLayers)
QUIT
N
"@
Set-Content -Path $extractScr -Value $extractScrContent -Encoding ASCII

Write-Host "PROGRESS: Scanning $($files.Count) files for layers..."

$allLayers = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

foreach ($file in $files) {
  Write-Host -NoNewline "."

  $base = Split-Path $file -LeafBase
  $outLog = Join-Path $ToolDir "extract_$base.out.txt"
  $errLog = Join-Path $ToolDir "extract_$base.err.txt"
  if (Test-Path $outLog) { Remove-Item $outLog -Force }
  if (Test-Path $errLog) { Remove-Item $errLog -Force }

  $result = Invoke-AcadCore -DwgPath $file -ScriptPath $extractScr -OutLog $outLog -ErrLog $errLog -TimeoutSeconds $ProcessTimeoutSeconds
  if ($result.TimedOut) { Write-Host "!" -NoNewline -ForegroundColor Red }
}
Write-Host "`nReading extracted data..."

if (Test-Path $layerDumpFile) {
  $rawLayers = Get-Content $layerDumpFile
  foreach ($line in $rawLayers) {
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    if ($line.Trim().StartsWith("###DWG:", [System.StringComparison]::OrdinalIgnoreCase)) { continue }
    [void]$allLayers.Add($line.Trim())
  }
}
else {
  Write-Warning "No layer dump file was created. Check logs in: $ToolDir"
}

if ($allLayers.Count -eq 0) {
  Write-Warning "0 layers found. Check: $ToolDir\extract_*.err.txt"
  exit 1
}

# ---------------- 5) USER SELECTION (GUI) ----------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Select Layers to FREEZE"
$form.Size = New-Object System.Drawing.Size(420, 540)
$form.StartPosition = "CenterScreen"

$lbl = New-Object System.Windows.Forms.Label
$lbl.Text = "Found $($allLayers.Count) unique layers:"
$lbl.Location = New-Object System.Drawing.Point(10, 10)
$lbl.Size = New-Object System.Drawing.Size(380, 20)
$form.Controls.Add($lbl)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10, 35)
$listBox.Size = New-Object System.Drawing.Size(380, 410)
$listBox.SelectionMode = "MultiExtended"

$sortedLayers = $allLayers | Sort-Object
$listBox.Items.AddRange([object[]]$sortedLayers)
$form.Controls.Add($listBox)

$btnOk = New-Object System.Windows.Forms.Button
$btnOk.Text = "Freeze Selected"
$btnOk.DialogResult = 'OK'
$btnOk.Location = New-Object System.Drawing.Point(10, 460)
$form.Controls.Add($btnOk)
$form.AcceptButton = $btnOk

if ($form.ShowDialog() -ne 'OK') { exit }

$layersToFreeze = @($listBox.SelectedItems)
if ($layersToFreeze.Count -eq 0) {
  Write-Host "No layers selected. Exiting."
  exit
}

# ---------------- 6) UPDATE PHASE (state-aware + report) ----------------
$freezeReport = Join-Path $ToolDir "FreezeReport.tsv"
"DWG`tLayer`tExists`tWasCurrent`tWasOff`tWasFrozen`tWasLocked`tAction" | Set-Content -Path $freezeReport -Encoding ASCII

$freezeReportForLisp = ($freezeReport -replace '\\', '/')

$updateLsp = Join-Path $ToolDir "update_layers.lsp"
$updateLspForLisp = ($updateLsp -replace '\\', '/')

# Build Lisp list: "Lay1" "Lay2"
$lispLayerList = $layersToFreeze | ForEach-Object { "`"$_`"" }
$lispLayerListStr = $lispLayerList -join " "

$updateLspContent = @"
(vl-load-com)

(defun _t (b) (if b "T" "F"))

(defun _log (f dwg lay exists wasCur wasOff wasFroz wasLock action / tab)
  (setq tab (chr 9))
  (write-line
    (strcat dwg tab lay tab exists tab wasCur tab wasOff tab wasFroz tab wasLock tab action)
    f
  )
)

(defun _safe-vla-item (layers name / r)
  (setq r (vl-catch-all-apply 'vla-item (list layers name)))
  (if (vl-catch-all-error-p r) nil r)
)

(defun _ensure-temp-current-layer (/ tmp)
  (setq tmp "HEADLESS_TEMP")
  ; Create if missing
  (if (null (tblsearch "LAYER" tmp))
    (command "_.-LAYER" "_New" tmp "")
  )
  ; Make current (also ensures it isn't frozen as current)
  (command "_.-LAYER" "_Make" tmp "")
  ; Make sure it's usable
  (command "_.-LAYER" "_On" tmp "")
  (command "_.-LAYER" "_Thaw" tmp "")
  (command "_.-LAYER" "_Unlock" tmp "")
  tmp
)

(defun c:BatchFreeze (/ doc layers f dwg layName exists layObj wasCur wasOff wasFroz wasLock action saveRes)

  (setvar "CMDECHO" 0)

  (setq f (open "$freezeReportForLisp" "a"))
  (setq doc (vla-get-activedocument (vlax-get-acad-object)))
  (setq layers (vla-get-layers doc))
  (setq dwg (getvar "DWGNAME"))

  (foreach layName (list $lispLayerListStr)

    (setq exists (if (tblsearch "LAYER" layName) T nil))
    (setq wasCur (= (strcase layName) (strcase (getvar "CLAYER"))))
    (setq wasOff nil)
    (setq wasFroz nil)
    (setq wasLock nil)
    (setq action "")

    (if (not exists)
      (progn
        (setq action "NOT_FOUND")
        (_log f dwg layName "F" (_t wasCur) "?" "?" "?" action)
      )
      (progn
        (setq layObj (_safe-vla-item layers layName))

        (if layObj
          (progn
            (setq wasOff  (= (vla-get-layeron layObj) :vlax-false))
            (setq wasFroz (= (vla-get-freeze  layObj) :vlax-true))
            (setq wasLock (= (vla-get-lock    layObj) :vlax-true))
          )
          (progn
            ; If we can't get the object (rare), still attempt freeze via command.
            (setq wasOff nil)
            (setq wasFroz nil)
            (setq wasLock nil)
          )
        )

        (cond
          (wasFroz
            (setq action "SKIP_ALREADY_FROZEN")
          )
          (T
            ; If it's current, switch CLAYER before freezing
            (if wasCur
              (progn
                (_ensure-temp-current-layer)
                (setq action (strcat action "SWITCH_CLAYER;"))
              )
            )

            ; If locked, unlock before freezing
            (if wasLock
              (progn
                (if layObj
                  (vl-catch-all-apply 'vla-put-lock (list layObj :vlax-false))
                  (command "_.-LAYER" "_Unlock" layName "")
                )
                (setq action (strcat action "UNLOCK;"))
              )
            )

            ; Freeze (do NOT change On/Off state)
            (if layObj
              (vl-catch-all-apply 'vla-put-freeze (list layObj :vlax-true))
              (command "_.-LAYER" "_Freeze" layName "")
            )
            (setq action (strcat action "FREEZE;"))

            (if wasOff
              (setq action (strcat action "LEFT_OFF;"))
            )
          )
        )

        (_log
          f dwg layName "T" (_t wasCur) (_t wasOff) (_t wasFroz) (_t wasLock) action
        )
      )
    )
  )

  ; Try to save (avoid QSAVE prompts)
  (setq saveRes (vl-catch-all-apply 'vla-save (list doc)))
  (if (vl-catch-all-error-p saveRes)
    (_log f dwg "<DWG_SAVE>" "T" "?" "?" "?" "?" (strcat "SAVE_FAILED:" (vl-catch-all-error-message saveRes)))
    (_log f dwg "<DWG_SAVE>" "T" "?" "?" "?" "?" "SAVE_OK")
  )

  (if f (close f))

  ; Exit without prompting (we already attempted save)
  (command "_.QUIT" "_N")
  (princ)
)
(princ)
"@
Set-Content -Path $updateLsp -Value $updateLspContent -Encoding ASCII

$updateScr = Join-Path $ToolDir "update.scr"
$updateScrContent = @"
FILEDIA
0
CMDDIA
0
PROXYNOTICE
0
SECURELOAD
0
(load "$updateLspForLisp")
(c:BatchFreeze)
"@
Set-Content -Path $updateScr -Value $updateScrContent -Encoding ASCII

# ---------------- 7) EXECUTE UPDATES ----------------
$logFile = Join-Path ([Environment]::GetFolderPath("Desktop")) "LayerUpdateLog.txt"
"Starting Update at $(Get-Date)" | Out-File $logFile

Write-Host "PROGRESS: Updating files..."

foreach ($file in $files) {
  $leaf = Split-Path $file -Leaf
  Write-Host "Processing: $leaf"

  $base = Split-Path $file -LeafBase
  $outLog = Join-Path $ToolDir "update_$base.out.txt"
  $errLog = Join-Path $ToolDir "update_$base.err.txt"
  if (Test-Path $outLog) { Remove-Item $outLog -Force }
  if (Test-Path $errLog) { Remove-Item $errLog -Force }

  $result = Invoke-AcadCore -DwgPath $file -ScriptPath $updateScr -OutLog $outLog -ErrLog $errLog -TimeoutSeconds $ProcessTimeoutSeconds

  if ($result.TimedOut) {
    Write-Host "  -> TIMEOUT. Killed." -ForegroundColor Red
    "FAILED (Timeout): $file" | Out-File $logFile -Append
    continue
  }

  if ($result.ExitCode -eq 0) {
    Write-Host "  -> CoreConsole ExitCode 0" -ForegroundColor Green
    "DONE (ExitCode 0): $file" | Out-File $logFile -Append
  }
  else {
    Write-Host "  -> Error Code: $($result.ExitCode)" -ForegroundColor Yellow
    "ERROR ($($result.ExitCode)): $file" | Out-File $logFile -Append
  }
}

Write-Host "Done."
Write-Host "Summary log: $logFile"
Write-Host "State/action report (per DWG + per layer): $freezeReport"
Write-Host "Per-file Core Console logs: $ToolDir"
