param(
  [string]$AcadCore
)

# ---------------- CONFIGURATION ----------------
$ProcessTimeoutSeconds = 180
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
    [int]$TimeoutSeconds = 180
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

  $r = Invoke-AcadCore -DwgPath $file -ScriptPath $extractScr -OutLog $outLog -ErrLog $errLog -TimeoutSeconds $ProcessTimeoutSeconds
  if ($r.TimedOut) { Write-Host "!" -NoNewline -ForegroundColor Red }
}
Write-Host "`nReading extracted data..."

if (-not (Test-Path $layerDumpFile)) {
  Write-Warning "No layer dump file was created. Check logs in: $ToolDir"
  exit 1
}

$rawLayers = Get-Content $layerDumpFile
foreach ($line in $rawLayers) {
  if ([string]::IsNullOrWhiteSpace($line)) { continue }
  if ($line.Trim().StartsWith("###DWG:", [System.StringComparison]::OrdinalIgnoreCase)) { continue }
  [void]$allLayers.Add($line.Trim())
}

if ($allLayers.Count -eq 0) {
  Write-Warning "0 layers found."
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
if ($layersToFreeze.Count -eq 0) { Write-Host "No layers selected. Exiting."; exit }

# ---------------- 6) UPDATE PHASE (STATE-AWARE, NO COM) ----------------
$freezeReport = Join-Path $ToolDir "FreezeReport.tsv"
"DWG`tLayer`tExists`tWasCurrent`tWasOff`tWasFrozen`tWasLocked`tAction`tResultFrozen`tSaveStatus" |
Set-Content -Path $freezeReport -Encoding ASCII

$freezeReportForLisp = ($freezeReport -replace '\\', '/')

$updateLsp = Join-Path $ToolDir "update_layers.lsp"
$updateLspForLisp = ($updateLsp -replace '\\', '/')

# Build Lisp list: "Lay1" "Lay2"
$lispLayerList = $layersToFreeze | ForEach-Object { "`"$_`"" }
$lispLayerListStr = $lispLayerList -join " "

$updateLspContent = @"
(defun _t (b) (if b "T" "F"))

(defun _flags (rec / v)
  (setq v (assoc 70 rec))
  (if v (cdr v) 0)
)

(defun _color62 (rec / v)
  (setq v (assoc 62 rec))
  (if v (cdr v) 0)
)

(defun _frozenp (rec)
  (/= 0 (logand (_flags rec) 1))
)

(defun _lockedp (rec)
  (/= 0 (logand (_flags rec) 4))
)

(defun _offp (rec)
  (< (_color62 rec) 0) ; group code 62 negative => off
)

(defun _log (f dwg lay exists wasCur wasOff wasFroz wasLock action resFroz saveStat / tab)
  (setq tab (chr 9))
  (write-line
    (strcat dwg tab lay tab exists tab wasCur tab wasOff tab wasFroz tab wasLock tab action tab resFroz tab saveStat)
    f
  )
)

(defun _ensure-temp-current-layer (/ tmp)
  (setq tmp "HEADLESS_TEMP")
  (if (null (tblsearch "LAYER" tmp))
    (command "_.-LAYER" "_New" tmp "")
  )
  (command "_.-LAYER" "_Thaw" tmp "")
  (command "_.-LAYER" "_On" tmp "")
  (command "_.-LAYER" "_Unlock" tmp "")
  (command "_.-LAYER" "_Make" tmp "")
  tmp
)

(defun c:BatchFreeze (/ f dwg saveBefore saveAfter saveStatus layName rec exists wasCur wasOff wasFroz wasLock action rec2 resFroz)

  (setvar "CMDECHO" 0)
  (setq dwg (getvar "DWGNAME"))

  (setq f (open "$freezeReportForLisp" "a"))
  (if (null f)
    (progn
      (prompt "\\nERROR: Could not open FreezeReport.tsv for append.")
      (command "_.QUIT" "_N")
      (princ)
    )
  )

  (foreach layName (list $lispLayerListStr)

    (setq rec (tblsearch "LAYER" layName))
    (setq exists (if rec T nil))

    (setq wasCur (if exists (= (strcase layName) (strcase (getvar "CLAYER"))) nil))
    (setq wasOff (if exists (_offp rec) nil))
    (setq wasFroz (if exists (_frozenp rec) nil))
    (setq wasLock (if exists (_lockedp rec) nil))
    (setq action "")

    (cond
      ((not exists)
        (setq action "NOT_FOUND")
        (setq resFroz "?")
      )
      (wasFroz
        ; Already frozen: do not alter anything
        (setq action "SKIP_ALREADY_FROZEN")
        (setq resFroz "T")
      )
      (T
        ; If current, switch to a safe layer first
        (if wasCur
          (progn (_ensure-temp-current-layer) (setq action (strcat action "SWITCH_CLAYER;")))
        )

        ; If locked, unlock then freeze
        (if wasLock
          (progn
            (command "_.-LAYER" "_Unlock" layName "")
            (setq action (strcat action "UNLOCK;"))
          )
        )

        ; Freeze (do not change ON/OFF)
        (command "_.-LAYER" "_Freeze" layName "")
        (setq action (strcat action "FREEZE;"))

        (if wasOff (setq action (strcat action "LEFT_OFF;")))

        ; Re-check frozen result
        (setq rec2 (tblsearch "LAYER" layName))
        (setq resFroz (if (and rec2 (_frozenp rec2)) "T" "F"))
      )
    )

    (_log f dwg layName
      (_t exists) (_t wasCur) (_t wasOff) (_t wasFroz) (_t wasLock)
      action resFroz "?"
    )
  )

  ; Save attempt + verification via DBMOD
  (setq saveBefore (getvar "DBMOD"))
  (command "_.QSAVE")
  (setq saveAfter (getvar "DBMOD"))

  (cond
    ((= saveAfter 0) (setq saveStatus "SAVE_OK"))
    (T (setq saveStatus (strcat "SAVE_FAILED_DBMOD=" (itoa saveAfter))))
  )

  ; Write one save status line
  (_log f dwg "<DWG_SAVE>" "T" "?" "?" "?" "?" "QSAVE" "?" saveStatus)

  (close f)

  ; Exit without further prompts (avoid hangs)
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

  $r = Invoke-AcadCore -DwgPath $file -ScriptPath $updateScr -OutLog $outLog -ErrLog $errLog -TimeoutSeconds $ProcessTimeoutSeconds

  if ($r.TimedOut) {
    Write-Host "  -> TIMEOUT. Killed." -ForegroundColor Red
    "FAILED (Timeout): $file" | Out-File $logFile -Append
    continue
  }

  if ($r.ExitCode -eq 0) {
    Write-Host "  -> ExitCode 0" -ForegroundColor Green
    "DONE (0): $file" | Out-File $logFile -Append
  }
  else {
    Write-Host "  -> ExitCode $($r.ExitCode)" -ForegroundColor Yellow
    "DONE ($($r.ExitCode)): $file" | Out-File $logFile -Append
  }
}

Write-Host "Done."
Write-Host "Report: $freezeReport"
Write-Host "Per-file logs: $ToolDir"
Write-Host "Summary log: $logFile"
