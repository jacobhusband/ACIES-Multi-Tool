param(
  [string]$AcadCore,
  [string]$ScanAllLayers = "true",
  [string]$FilesListPath = "",
  [string]$DefaultDirectory = ""
)

function Convert-ToBool {
  param(
    [object]$Value,
    [bool]$DefaultValue = $true
  )
  if ($null -eq $Value) { return $DefaultValue }
  $text = $Value.ToString().Trim()
  if ($text.StartsWith('$')) { $text = $text.Substring(1) }
  switch -Regex ($text) {
    '^(1|true|t|yes|y)$' { return $true }
    '^(0|false|f|no|n)$' { return $false }
    default { return $DefaultValue }
  }
}

$ScanAllLayers = Convert-ToBool $ScanAllLayers $true

function Resolve-DialogInitialDirectory {
  param(
    [string]$CandidatePath,
    [string]$FallbackPath = ""
  )

  if ([string]::IsNullOrWhiteSpace($CandidatePath)) { return $FallbackPath }
  $resolvedPath = $CandidatePath.Trim()
  if (Test-Path -Path $resolvedPath -PathType Leaf) {
    $resolvedPath = Split-Path -Path $resolvedPath -Parent
  }
  if (-not [string]::IsNullOrWhiteSpace($resolvedPath) -and (Test-Path -Path $resolvedPath -PathType Container)) {
    return $resolvedPath
  }
  return $FallbackPath
}

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
  $argsList = @("-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-File", $PSCommandPath)
  if ($AcadCore) { $argsList += @("-AcadCore", $AcadCore) }
  if ($PSBoundParameters.ContainsKey('ScanAllLayers')) {
    $argsList += @("-ScanAllLayers", $ScanAllLayers)
  }
  if ($PSBoundParameters.ContainsKey('FilesListPath') -and -not [string]::IsNullOrWhiteSpace($FilesListPath)) {
    $argsList += @("-FilesListPath", $FilesListPath)
  }
  if ($PSBoundParameters.ContainsKey('DefaultDirectory') -and -not [string]::IsNullOrWhiteSpace($DefaultDirectory)) {
    $argsList += @("-DefaultDirectory", $DefaultDirectory)
  }
  $child = Start-Process -FilePath $ps -ArgumentList $argsList -Wait -PassThru
  exit $child.ExitCode
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Show-DwgFileSelectionPrompt {
  $promptForm = New-Object System.Windows.Forms.Form
  $promptForm.Text = "Select DWG file(s)"
  $promptForm.StartPosition = "CenterScreen"
  $promptForm.Size = New-Object System.Drawing.Size(560, 190)
  $promptForm.MinimumSize = New-Object System.Drawing.Size(560, 190)
  $promptForm.MaximumSize = New-Object System.Drawing.Size(560, 190)
  $promptForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
  $promptForm.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
  $promptForm.MaximizeBox = $false
  $promptForm.MinimizeBox = $false
  $promptForm.TopMost = $true
  $promptForm.ShowInTaskbar = $true

  $lblPrompt = New-Object System.Windows.Forms.Label
  $lblPrompt.Text = "Choose one or more DWG files to process. This window stays on top until files are selected or you exit."
  $lblPrompt.Location = New-Object System.Drawing.Point(16, 16)
  $lblPrompt.Size = New-Object System.Drawing.Size(520, 52)
  $lblPrompt.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
  $promptForm.Controls.Add($lblPrompt)

  $btnSelectFiles = New-Object System.Windows.Forms.Button
  $btnSelectFiles.Text = "Select DWG Files..."
  $btnSelectFiles.Size = New-Object System.Drawing.Size(210, 44)
  $btnSelectFiles.Location = New-Object System.Drawing.Point(222, 92)
  $btnSelectFiles.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
  $btnSelectFiles.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
  $btnSelectFiles.FlatStyle = [System.Windows.Forms.FlatStyle]::System
  $promptForm.Controls.Add($btnSelectFiles)

  $btnExitPrompt = New-Object System.Windows.Forms.Button
  $btnExitPrompt.Text = "Exit"
  $btnExitPrompt.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
  $btnExitPrompt.Size = New-Object System.Drawing.Size(110, 44)
  $btnExitPrompt.Location = New-Object System.Drawing.Point(434, 92)
  $btnExitPrompt.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
  $btnExitPrompt.Font = New-Object System.Drawing.Font("Segoe UI", 10)
  $btnExitPrompt.FlatStyle = [System.Windows.Forms.FlatStyle]::System
  $promptForm.Controls.Add($btnExitPrompt)

  $promptForm.CancelButton = $btnExitPrompt
  $promptForm.AcceptButton = $btnSelectFiles

  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Title = "Select DWG file(s)"
  $dlg.Filter = "DWG files (*.dwg)|*.dwg"
  $dlg.Multiselect = $true
  $dlg.CheckFileExists = $true
  $dlg.CheckPathExists = $true
  $dlg.RestoreDirectory = $true
  $dlg.InitialDirectory = Resolve-DialogInitialDirectory -CandidatePath $DefaultDirectory

  $btnSelectFiles.add_Click({
      $promptForm.TopMost = $true
      $promptForm.Activate()
      $result = $dlg.ShowDialog($promptForm)
      if ($result -eq [System.Windows.Forms.DialogResult]::OK -and $dlg.FileNames.Count -gt 0) {
        $promptForm.Tag = [string[]]$dlg.FileNames
        $promptForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $promptForm.Close()
        return
      }
      $promptForm.TopMost = $true
      $promptForm.Activate()
      $promptForm.BringToFront()
    })

  $promptForm.add_Shown({
      $promptForm.Activate()
      $promptForm.BringToFront()
      $btnSelectFiles.PerformClick()
    })

  # Keep this prompt from being minimized so it remains actionable.
  $promptForm.add_Resize({
      if ($promptForm.WindowState -eq [System.Windows.Forms.FormWindowState]::Minimized) {
        $promptForm.WindowState = [System.Windows.Forms.FormWindowState]::Normal
        $promptForm.Activate()
        $promptForm.BringToFront()
      }
    })

  $promptResult = $promptForm.ShowDialog()
  if ($promptResult -eq [System.Windows.Forms.DialogResult]::OK) {
    $selectedFiles = @($promptForm.Tag)
    if ($selectedFiles.Count -gt 0) {
      return $selectedFiles
    }
  }
  return $null
}

function Move-FormToPrimaryScreen {
  param([System.Windows.Forms.Form]$TargetForm)

  if ($null -eq $TargetForm) { return }
  $workingArea = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
  $x = $workingArea.Left + [Math]::Max(0, [int](($workingArea.Width - $TargetForm.Width) / 2))
  $y = $workingArea.Top + [Math]::Max(0, [int](($workingArea.Height - $TargetForm.Height) / 2))
  $TargetForm.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
  $TargetForm.Location = New-Object System.Drawing.Point($x, $y)
}

$files = @()
$filesListWasProvided = $PSBoundParameters.ContainsKey('FilesListPath')
$hasFilesListPath = $filesListWasProvided -and -not [string]::IsNullOrWhiteSpace($FilesListPath)
Write-Host "PROGRESS: TRACE files_list_param_bound=$([int]$filesListWasProvided) path=$FilesListPath"
if ($filesListWasProvided) {
  if ($hasFilesListPath) {
    Write-Host "PROGRESS: Received auto-selected files list: $FilesListPath"
    if (Test-Path $FilesListPath) {
      Write-Host "PROGRESS: TRACE files_list_path_exists=1"
      $files = @(
        Get-Content -Path $FilesListPath -Encoding UTF8 |
          Where-Object { $_ -and $_.Trim() -and (Test-Path $_.Trim()) } |
          ForEach-Object { $_.Trim() }
      )
      Write-Host "PROGRESS: TRACE files_list_valid_count=$($files.Count)"
      if ($files.Count -gt 0) {
        Write-Host "PROGRESS: Using $($files.Count) DWG file(s) from auto-selected project folder."
      }
      else {
        Write-Host "PROGRESS: Provided files list was empty. Opening file picker..."
      }
    }
    else {
      Write-Host "PROGRESS: TRACE files_list_path_exists=0"
      Write-Host "PROGRESS: Provided files list path was not found. Opening file picker..."
    }
  }
  else {
    Write-Host "PROGRESS: TRACE files_list_param_empty=1"
    Write-Host "PROGRESS: Files list parameter was provided without a path. Opening file picker..."
  }
}

if ($files -and $files.Count -gt 0) {
  Write-Host "PROGRESS: TRACE branch=auto_selected_files count=$($files.Count)"
}

if (-not $files -or $files.Count -eq 0) {
  Write-Host "PROGRESS: TRACE branch=manual_picker"
  Write-Host "PROGRESS: Waiting for user input..."
  $files = Show-DwgFileSelectionPrompt
  if (-not $files -or $files.Count -eq 0) { exit }
}

# Normalize to a string array so single-file runs still behave like multi-file runs.
$files = @($files)

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
$scanListFile = Join-Path $ToolDir "scan_files.txt"
$scanListForLisp = ($scanListFile -replace '\\', '/')

$extractLspContent = @"
(vl-load-com)

(defun _read-lines (path / f line lines)
  (setq f (open path "r"))
  (if f
    (progn
      (while (setq line (read-line f))
        (if (> (strlen line) 0)
          (setq lines (cons line lines))
        )
      )
      (close f)
      (reverse lines)
    )
  )
)

;;; Convert forward slashes to backslashes for Windows paths
(defun _fix-path (path)
  (vl-string-translate "/" "\\" path)
)

;;; Fallback: Original tblnext-based extraction for current document
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

;;; Try to create ObjectDBX document with version-specific ProgIDs
(defun _create-dbx-doc (/ dbxDoc progIds progId)
  (setq progIds (list
    "ObjectDBX.AxDbDocument.25"  ; AutoCAD 2026
    "ObjectDBX.AxDbDocument.24"  ; AutoCAD 2025
    "ObjectDBX.AxDbDocument.23"  ; AutoCAD 2024
    "ObjectDBX.AxDbDocument.22"  ; AutoCAD 2023
    "ObjectDBX.AxDbDocument.21"  ; AutoCAD 2022
    "ObjectDBX.AxDbDocument.20"  ; AutoCAD 2021
    "ObjectDBX.AxDbDocument.19"  ; AutoCAD 2020
    "ObjectDBX.AxDbDocument.18"  ; AutoCAD 2019
    "ObjectDBX.AxDbDocument"     ; Generic fallback
  ))
  (setq dbxDoc nil)
  (foreach progId progIds
    (if (null dbxDoc)
      (setq dbxDoc (vl-catch-all-apply 'vlax-create-object (list progId)))
    )
    (if (vl-catch-all-error-p dbxDoc)
      (setq dbxDoc nil)
    )
  )
  dbxDoc
)

;;; ObjectDBX-based layer extraction (fast, database-only access)
(defun _extract-layers-dbx (dwgPath reportFile / fixedPath dbxDoc layers layObj f openErr)
  (setq fixedPath (_fix-path dwgPath))
  (if fixedPath
    (progn
      (setq dbxDoc (_create-dbx-doc))
      (if dbxDoc
        (progn
          (setq openErr
            (vl-catch-all-apply
              'vlax-invoke-method
              (list dbxDoc 'Open fixedPath)))

          (if (not (vl-catch-all-error-p openErr))
            (progn
              (setq f (open reportFile "a"))
              (if f
                (progn
                  (write-line
                    (strcat "###DWG:" (strcat (vl-filename-base fixedPath) ".dwg"))
                    f)
                  (setq layers (vlax-get-property dbxDoc 'Layers))
                  (vlax-for layObj layers
                    (write-line (vlax-get-property layObj 'Name) f)
                  )
                  (close f)
                  (vlax-release-object dbxDoc)
                  T
                )
                (progn (vlax-release-object dbxDoc) nil)
              )
            )
            (progn (vlax-release-object dbxDoc) nil)
          )
        )
        nil
      )
    )
    nil
  )
)

;;; Batch: Try ObjectDBX first, fall back to OPEN (with xrefs disabled) for failures
(defun c:ExtractLayersBatch (/ files dwgPath failedFiles oldXloadctl)
  (setq files (_read-lines "$scanListForLisp"))
  (setq failedFiles nil)

  (if files
    (progn
      ;; First pass: Try ObjectDBX for all files
      (foreach dwgPath files
        (if (not (_extract-layers-dbx dwgPath "$lispReportPath"))
          (setq failedFiles (cons dwgPath failedFiles))
        )
      )

      ;; Second pass: Fallback with XLOADCTL=0 to skip xref loading
      (if failedFiles
        (progn
          (setq oldXloadctl (getvar "XLOADCTL"))
          (setvar "XLOADCTL" 0)
          (foreach f (reverse failedFiles)
            (command "_.OPEN" f)
            (c:ExtractLayers)
          )
          (setvar "XLOADCTL" oldXloadctl)
        )
      )
    )
  )
  (command "_.QUIT" "_N")
  (princ)
)
(princ)
"@
Set-Content -Path $extractLsp -Value $extractLspContent -Encoding ASCII

$extractScr = Join-Path $ToolDir "extract.scr"

if ($ScanAllLayers) {
  Write-Host "PROGRESS: Scanning $($files.Count) files for layers..."
}
else {
  Write-Host "PROGRESS: Scanning first file only for layers..."
}
$allLayers = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

$filesToScan = if ($ScanAllLayers) { $files } else { @($files[0]) }
if ($filesToScan.Count -eq 0) {
  Write-Warning "No files selected."
  exit
}

$filesToScan | ForEach-Object { ($_ -replace '\\', '/') } |
  Set-Content -Path $scanListFile -Encoding ASCII

$extractScrLines = @(
  "FILEDIA",
  "0",
  "CMDDIA",
  "0",
  "PROXYNOTICE",
  "0",
  "SECURELOAD",
  "0",
  "(load `"$extractLspForLisp`")",
  "(c:ExtractLayersBatch)"
)
Set-Content -Path $extractScr -Value ($extractScrLines -join "`r`n") -Encoding ASCII

$outLog = Join-Path $ToolDir "extract_batch.out.txt"
$errLog = Join-Path $ToolDir "extract_batch.err.txt"
if (Test-Path $outLog) { Remove-Item $outLog -Force }
if (Test-Path $errLog) { Remove-Item $errLog -Force }

$scanTimeoutSeconds = $ProcessTimeoutSeconds * $filesToScan.Count
$r = Invoke-AcadCore -DwgPath $filesToScan[0] -ScriptPath $extractScr -OutLog $outLog -ErrLog $errLog -TimeoutSeconds $scanTimeoutSeconds
if ($r.TimedOut) { Write-Host "Layer scan timed out." -ForegroundColor Red }

Write-Host "PROGRESS: Reading extracted data..."

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

# ---------------- 5) USER SELECTION (GUI with THREE LISTS) ----------------
$allSorted = @($allLayers | Sort-Object)
$layersToFreezeSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$layersToThawSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

$form = New-Object System.Windows.Forms.Form
$form.Text = "Select Layers to Freeze or Thaw"
$form.StartPosition = "Manual"
$form.Size = New-Object System.Drawing.Size(1120, 740)
$form.MinimumSize = New-Object System.Drawing.Size(1120, 740)
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
$form.TopMost = $true
$form.ShowInTaskbar = $true
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
Move-FormToPrimaryScreen $form

$lbl = New-Object System.Windows.Forms.Label
$lbl.Location = New-Object System.Drawing.Point(12, 12)
$lbl.Size = New-Object System.Drawing.Size(1090, 20)
$lbl.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($lbl)

$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Text = "Filter available:"
$lblFilter.Location = New-Object System.Drawing.Point(12, 40)
$lblFilter.Size = New-Object System.Drawing.Size(98, 22)
$form.Controls.Add($lblFilter)

$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.Location = New-Object System.Drawing.Point(112, 38)
$txtFilter.Size = New-Object System.Drawing.Size(240, 26)
$txtFilter.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($txtFilter)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "Clear"
$btnClear.Location = New-Object System.Drawing.Point(360, 36)
$btnClear.Size = New-Object System.Drawing.Size(70, 28)
$btnClear.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$btnClear.FlatStyle = [System.Windows.Forms.FlatStyle]::System
$form.Controls.Add($btnClear)

$lblAvailable = New-Object System.Windows.Forms.Label
$lblAvailable.Location = New-Object System.Drawing.Point(12, 74)
$lblAvailable.Size = New-Object System.Drawing.Size(400, 18)
$lblAvailable.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($lblAvailable)

$lblFreeze = New-Object System.Windows.Forms.Label
$lblFreeze.Location = New-Object System.Drawing.Point(538, 74)
$lblFreeze.Size = New-Object System.Drawing.Size(540, 18)
$lblFreeze.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($lblFreeze)

$lblThaw = New-Object System.Windows.Forms.Label
$lblThaw.Location = New-Object System.Drawing.Point(538, 338)
$lblThaw.Size = New-Object System.Drawing.Size(540, 18)
$lblThaw.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($lblThaw)

$listAvailable = New-Object System.Windows.Forms.ListBox
$listAvailable.Location = New-Object System.Drawing.Point(12, 96)
$listAvailable.Size = New-Object System.Drawing.Size(400, 504)
$listAvailable.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$listAvailable.SelectionMode = "MultiExtended"
$listAvailable.IntegralHeight = $false
$form.Controls.Add($listAvailable)

# --- Freeze button panel (middle-upper) ---
$movePanelFreeze = New-Object System.Windows.Forms.Panel
$movePanelFreeze.Location = New-Object System.Drawing.Point(418, 120)
$movePanelFreeze.Size = New-Object System.Drawing.Size(114, 196)
$movePanelFreeze.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($movePanelFreeze)

$btnAddFreeze = New-Object System.Windows.Forms.Button
$btnAddFreeze.Text = ">"
$btnAddFreeze.Size = New-Object System.Drawing.Size(80, 38)
$btnAddFreeze.Location = New-Object System.Drawing.Point(16, 14)
$btnAddFreeze.FlatStyle = [System.Windows.Forms.FlatStyle]::System
$movePanelFreeze.Controls.Add($btnAddFreeze)

$btnAddAllFreeze = New-Object System.Windows.Forms.Button
$btnAddAllFreeze.Text = ">>"
$btnAddAllFreeze.Size = New-Object System.Drawing.Size(80, 38)
$btnAddAllFreeze.Location = New-Object System.Drawing.Point(16, 58)
$btnAddAllFreeze.FlatStyle = [System.Windows.Forms.FlatStyle]::System
$movePanelFreeze.Controls.Add($btnAddAllFreeze)

$btnRemoveFreeze = New-Object System.Windows.Forms.Button
$btnRemoveFreeze.Text = "<"
$btnRemoveFreeze.Size = New-Object System.Drawing.Size(80, 38)
$btnRemoveFreeze.Location = New-Object System.Drawing.Point(16, 108)
$btnRemoveFreeze.FlatStyle = [System.Windows.Forms.FlatStyle]::System
$movePanelFreeze.Controls.Add($btnRemoveFreeze)

$btnRemoveAllFreeze = New-Object System.Windows.Forms.Button
$btnRemoveAllFreeze.Text = "<<"
$btnRemoveAllFreeze.Size = New-Object System.Drawing.Size(80, 38)
$btnRemoveAllFreeze.Location = New-Object System.Drawing.Point(16, 152)
$btnRemoveAllFreeze.FlatStyle = [System.Windows.Forms.FlatStyle]::System
$movePanelFreeze.Controls.Add($btnRemoveAllFreeze)

# --- Thaw button panel (middle-lower) ---
$movePanelThaw = New-Object System.Windows.Forms.Panel
$movePanelThaw.Location = New-Object System.Drawing.Point(418, 384)
$movePanelThaw.Size = New-Object System.Drawing.Size(114, 196)
$movePanelThaw.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($movePanelThaw)

$btnAddThaw = New-Object System.Windows.Forms.Button
$btnAddThaw.Text = ">"
$btnAddThaw.Size = New-Object System.Drawing.Size(80, 38)
$btnAddThaw.Location = New-Object System.Drawing.Point(16, 14)
$btnAddThaw.FlatStyle = [System.Windows.Forms.FlatStyle]::System
$movePanelThaw.Controls.Add($btnAddThaw)

$btnAddAllThaw = New-Object System.Windows.Forms.Button
$btnAddAllThaw.Text = ">>"
$btnAddAllThaw.Size = New-Object System.Drawing.Size(80, 38)
$btnAddAllThaw.Location = New-Object System.Drawing.Point(16, 58)
$btnAddAllThaw.FlatStyle = [System.Windows.Forms.FlatStyle]::System
$movePanelThaw.Controls.Add($btnAddAllThaw)

$btnRemoveThaw = New-Object System.Windows.Forms.Button
$btnRemoveThaw.Text = "<"
$btnRemoveThaw.Size = New-Object System.Drawing.Size(80, 38)
$btnRemoveThaw.Location = New-Object System.Drawing.Point(16, 108)
$btnRemoveThaw.FlatStyle = [System.Windows.Forms.FlatStyle]::System
$movePanelThaw.Controls.Add($btnRemoveThaw)

$btnRemoveAllThaw = New-Object System.Windows.Forms.Button
$btnRemoveAllThaw.Text = "<<"
$btnRemoveAllThaw.Size = New-Object System.Drawing.Size(80, 38)
$btnRemoveAllThaw.Location = New-Object System.Drawing.Point(16, 152)
$btnRemoveAllThaw.FlatStyle = [System.Windows.Forms.FlatStyle]::System
$movePanelThaw.Controls.Add($btnRemoveAllThaw)

$listToFreeze = New-Object System.Windows.Forms.ListBox
$listToFreeze.Location = New-Object System.Drawing.Point(538, 96)
$listToFreeze.Size = New-Object System.Drawing.Size(540, 236)
$listToFreeze.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$listToFreeze.SelectionMode = "MultiExtended"
$listToFreeze.IntegralHeight = $false
$form.Controls.Add($listToFreeze)

$listToThaw = New-Object System.Windows.Forms.ListBox
$listToThaw.Location = New-Object System.Drawing.Point(538, 360)
$listToThaw.Size = New-Object System.Drawing.Size(540, 240)
$listToThaw.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$listToThaw.SelectionMode = "MultiExtended"
$listToThaw.IntegralHeight = $false
$form.Controls.Add($listToThaw)

$btnOk = New-Object System.Windows.Forms.Button
$btnOk.Text = "Apply Freeze + Thaw"
$btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
$btnOk.Size = New-Object System.Drawing.Size(240, 48)
$btnOk.Location = New-Object System.Drawing.Point(838, 620)
$btnOk.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$btnOk.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnOk.FlatStyle = [System.Windows.Forms.FlatStyle]::System
$form.Controls.Add($btnOk)
$form.AcceptButton = $btnOk

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel"
$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$btnCancel.Size = New-Object System.Drawing.Size(110, 48)
$btnCancel.Location = New-Object System.Drawing.Point(720, 620)
$btnCancel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$btnCancel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::System
$form.Controls.Add($btnCancel)
$form.CancelButton = $btnCancel

$form.add_Shown({
    $form.TopMost = $true
    $form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
    Move-FormToPrimaryScreen $form
    $form.Activate()
    $form.BringToFront()
    $listAvailable.Focus()
  })

$form.add_Resize({
    if ($form.WindowState -eq [System.Windows.Forms.FormWindowState]::Minimized) {
      $form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
      $form.Activate()
      $form.BringToFront()
    }
  })

function Update-MoveButtons {
  $btnAddFreeze.Enabled = $listAvailable.SelectedItems.Count -gt 0
  $btnAddAllFreeze.Enabled = $listAvailable.Items.Count -gt 0
  $btnRemoveFreeze.Enabled = $listToFreeze.SelectedItems.Count -gt 0
  $btnRemoveAllFreeze.Enabled = $listToFreeze.Items.Count -gt 0

  $btnAddThaw.Enabled = $listAvailable.SelectedItems.Count -gt 0
  $btnAddAllThaw.Enabled = $listAvailable.Items.Count -gt 0
  $btnRemoveThaw.Enabled = $listToThaw.SelectedItems.Count -gt 0
  $btnRemoveAllThaw.Enabled = $listToThaw.Items.Count -gt 0
}

function Refresh-LayerLists {
  param([string]$filterText)

  if ($null -eq $filterText) { $filterText = "" }
  $filterText = $filterText.Trim()

  $selectedAvailable = @()
  foreach ($item in $listAvailable.SelectedItems) { $selectedAvailable += $item.ToString() }
  $selectedToFreeze = @()
  foreach ($item in $listToFreeze.SelectedItems) { $selectedToFreeze += $item.ToString() }
  $selectedToThaw = @()
  foreach ($item in $listToThaw.SelectedItems) { $selectedToThaw += $item.ToString() }

  $availableLayers = @($allSorted | Where-Object {
      -not $layersToFreezeSet.Contains($_) -and -not $layersToThawSet.Contains($_)
    })
  if (-not [string]::IsNullOrWhiteSpace($filterText)) {
    $availableLayers = @($availableLayers | Where-Object { $_.IndexOf($filterText, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 })
  }

  $layersToFreezeSorted = @($layersToFreezeSet | Sort-Object)
  $layersToThawSorted = @($layersToThawSet | Sort-Object)

  $listAvailable.BeginUpdate()
  $listAvailable.Items.Clear()
  if ($availableLayers.Count -gt 0) {
    $listAvailable.Items.AddRange([object[]]$availableLayers)
  }
  $listAvailable.EndUpdate()

  $listToFreeze.BeginUpdate()
  $listToFreeze.Items.Clear()
  if ($layersToFreezeSorted.Count -gt 0) {
    $listToFreeze.Items.AddRange([object[]]$layersToFreezeSorted)
  }
  $listToFreeze.EndUpdate()

  $listToThaw.BeginUpdate()
  $listToThaw.Items.Clear()
  if ($layersToThawSorted.Count -gt 0) {
    $listToThaw.Items.AddRange([object[]]$layersToThawSorted)
  }
  $listToThaw.EndUpdate()

  foreach ($item in $selectedAvailable) {
    $idx = $listAvailable.Items.IndexOf($item)
    if ($idx -ge 0) { $listAvailable.SetSelected($idx, $true) }
  }
  foreach ($item in $selectedToFreeze) {
    $idx = $listToFreeze.Items.IndexOf($item)
    if ($idx -ge 0) { $listToFreeze.SetSelected($idx, $true) }
  }
  foreach ($item in $selectedToThaw) {
    $idx = $listToThaw.Items.IndexOf($item)
    if ($idx -ge 0) { $listToThaw.SetSelected($idx, $true) }
  }

  $lbl.Text = "Move layers to the Freeze list to freeze them, or to the Thaw list to thaw them. Layers left on the left are untouched."
  $lblAvailable.Text = "Available: $($availableLayers.Count)"
  $lblFreeze.Text = "To Freeze: $($layersToFreezeSorted.Count)"
  $lblThaw.Text = "To Thaw: $($layersToThawSorted.Count)"
  $btnOk.Enabled = ($layersToFreezeSorted.Count + $layersToThawSorted.Count) -gt 0
  Update-MoveButtons
}

function Add-SelectedAvailableToFreeze {
  $selected = @()
  foreach ($item in $listAvailable.SelectedItems) { $selected += $item.ToString() }
  foreach ($layerName in $selected) {
    [void]$layersToThawSet.Remove($layerName)
    [void]$layersToFreezeSet.Add($layerName)
  }
  Refresh-LayerLists $txtFilter.Text
  $listAvailable.Focus()
}

function Add-AllVisibleAvailableToFreeze {
  $visible = @()
  foreach ($item in $listAvailable.Items) { $visible += $item.ToString() }
  foreach ($layerName in $visible) {
    [void]$layersToThawSet.Remove($layerName)
    [void]$layersToFreezeSet.Add($layerName)
  }
  Refresh-LayerLists $txtFilter.Text
  $listAvailable.Focus()
}

function Remove-SelectedFreezeLayers {
  $selected = @()
  foreach ($item in $listToFreeze.SelectedItems) { $selected += $item.ToString() }
  foreach ($layerName in $selected) { [void]$layersToFreezeSet.Remove($layerName) }
  Refresh-LayerLists $txtFilter.Text
  $listToFreeze.Focus()
}

function Remove-AllFreezeLayers {
  $layersToFreezeSet.Clear()
  Refresh-LayerLists $txtFilter.Text
  $listAvailable.Focus()
}

function Add-SelectedAvailableToThaw {
  $selected = @()
  foreach ($item in $listAvailable.SelectedItems) { $selected += $item.ToString() }
  foreach ($layerName in $selected) {
    [void]$layersToFreezeSet.Remove($layerName)
    [void]$layersToThawSet.Add($layerName)
  }
  Refresh-LayerLists $txtFilter.Text
  $listAvailable.Focus()
}

function Add-AllVisibleAvailableToThaw {
  $visible = @()
  foreach ($item in $listAvailable.Items) { $visible += $item.ToString() }
  foreach ($layerName in $visible) {
    [void]$layersToFreezeSet.Remove($layerName)
    [void]$layersToThawSet.Add($layerName)
  }
  Refresh-LayerLists $txtFilter.Text
  $listAvailable.Focus()
}

function Remove-SelectedThawLayers {
  $selected = @()
  foreach ($item in $listToThaw.SelectedItems) { $selected += $item.ToString() }
  foreach ($layerName in $selected) { [void]$layersToThawSet.Remove($layerName) }
  Refresh-LayerLists $txtFilter.Text
  $listToThaw.Focus()
}

function Remove-AllThawLayers {
  $layersToThawSet.Clear()
  Refresh-LayerLists $txtFilter.Text
  $listAvailable.Focus()
}

$txtFilter.add_TextChanged({ Refresh-LayerLists $txtFilter.Text })

$btnClear.add_Click({
    $txtFilter.Text = ""
    $txtFilter.Focus()
  })

$listAvailable.add_SelectedIndexChanged({ Update-MoveButtons })
$listToFreeze.add_SelectedIndexChanged({ Update-MoveButtons })
$listToThaw.add_SelectedIndexChanged({ Update-MoveButtons })

$btnAddFreeze.add_Click({ Add-SelectedAvailableToFreeze })
$btnAddAllFreeze.add_Click({ Add-AllVisibleAvailableToFreeze })
$btnRemoveFreeze.add_Click({ Remove-SelectedFreezeLayers })
$btnRemoveAllFreeze.add_Click({ Remove-AllFreezeLayers })

$btnAddThaw.add_Click({ Add-SelectedAvailableToThaw })
$btnAddAllThaw.add_Click({ Add-AllVisibleAvailableToThaw })
$btnRemoveThaw.add_Click({ Remove-SelectedThawLayers })
$btnRemoveAllThaw.add_Click({ Remove-AllThawLayers })

$listToFreeze.add_DoubleClick({ Remove-SelectedFreezeLayers })
$listToThaw.add_DoubleClick({ Remove-SelectedThawLayers })

$listToFreeze.add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Delete -or $e.KeyCode -eq [System.Windows.Forms.Keys]::Back) {
      Remove-SelectedFreezeLayers
      $e.Handled = $true
    }
  })

$listToThaw.add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Delete -or $e.KeyCode -eq [System.Windows.Forms.Keys]::Back) {
      Remove-SelectedThawLayers
      $e.Handled = $true
    }
  })

Refresh-LayerLists ""

Write-Host "PROGRESS: Waiting for layer selection..."
Write-Host "PROGRESS: Layer selection dialog should be visible on the primary display."
Write-Host "PROGRESS: TRACE branch=layer_selection_dialog"
if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { exit }

$layersToFreeze = @($layersToFreezeSet | Sort-Object)
$layersToThaw = @($layersToThawSet | Sort-Object)
if ($layersToFreeze.Count -eq 0 -and $layersToThaw.Count -eq 0) {
  Write-Host "No layers were queued for freeze or thaw. Exiting."
  exit
}

# ---------------- 6) UPDATE PHASE (STATE-AWARE, NO COM) ----------------
$updateReport = Join-Path $ToolDir "LayerUpdateReport.tsv"
"DWG`tLayer`tOpType`tExists`tWasCurrent`tWasOff`tWasFrozen`tWasLocked`tAction`tResultFrozen`tSaveStatus" |
  Set-Content -Path $updateReport -Encoding ASCII

$updateReportForLisp = ($updateReport -replace '\\', '/')

$updateLsp = Join-Path $ToolDir "update_layers.lsp"
$updateLspForLisp = ($updateLsp -replace '\\', '/')
$lispFreezeList = if ($layersToFreeze.Count -gt 0) { ($layersToFreeze | ForEach-Object { "`"$_`"" }) -join " " } else { "" }
$lispThawList = if ($layersToThaw.Count -gt 0) { ($layersToThaw | ForEach-Object { "`"$_`"" }) -join " " } else { "" }

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
  (< (_color62 rec) 0)
)

(defun _log (f dwg lay opType exists wasCur wasOff wasFroz wasLock action resFroz saveStat / tab)
  (setq tab (chr 9))
  (write-line
    (strcat dwg tab lay tab opType tab exists tab wasCur tab wasOff tab wasFroz tab wasLock tab action tab resFroz tab saveStat)
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

(defun _freeze-one (f dwg layName freezeSet / rec exists wasCur wasOff wasFroz wasLock action rec2 resFroz)
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
      (setq action "SKIP_ALREADY_FROZEN")
      (setq resFroz "T")
    )
    (T
      (if wasCur (progn (_ensure-temp-current-layer) (setq action (strcat action "SWITCH_CLAYER;"))))
      (if wasLock (progn (command "_.-LAYER" "_Unlock" layName "") (setq action (strcat action "UNLOCK;"))))

      (command "_.-LAYER" "_Freeze" layName "")
      (setq action (strcat action "FREEZE;"))
      (if wasOff (setq action (strcat action "LEFT_OFF;")))

      (setq rec2 (tblsearch "LAYER" layName))
      (setq resFroz (if (and rec2 (_frozenp rec2)) "T" "F"))
    )
  )

  (_log f dwg layName "FREEZE"
    (_t exists) (_t wasCur) (_t wasOff) (_t wasFroz) (_t wasLock)
    action resFroz "?"
  )
)

(defun _thaw-one (f dwg layName freezeSet / rec exists wasCur wasOff wasFroz wasLock action rec2 resFroz)
  ;; Defensive: if this layer is also queued for freeze, prefer freeze and skip thaw.
  (if (member (strcase layName) freezeSet)
    (progn
      (_log f dwg layName "THAW" "?" "?" "?" "?" "?" "CONFLICT_FREEZE_WINS" "?" "?")
    )
    (progn
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
        ((not wasFroz)
          (setq action "SKIP_ALREADY_THAWED")
          (setq resFroz "F")
        )
        (T
          (command "_.-LAYER" "_Thaw" layName "")
          (setq action "THAW;")
          (setq rec2 (tblsearch "LAYER" layName))
          (setq resFroz (if (and rec2 (_frozenp rec2)) "T" "F"))
        )
      )

      (_log f dwg layName "THAW"
        (_t exists) (_t wasCur) (_t wasOff) (_t wasFroz) (_t wasLock)
        action resFroz "?"
      )
    )
  )
)

(defun c:BatchUpdate (/ f dwg freezeList thawList freezeSet saveAfter saveStatus)

  (setvar "CMDECHO" 0)
  (setq dwg (getvar "DWGNAME"))

  (setq f (open "$updateReportForLisp" "a"))
  (if (null f)
    (progn
      (prompt "\\nERROR: Could not open LayerUpdateReport.tsv for append.")
      (command "_.QUIT" "_N")
      (princ)
    )
  )

  (setq freezeList (list $lispFreezeList))
  (setq thawList (list $lispThawList))

  ;; Filter out empty placeholder from empty lists.
  (setq freezeList (vl-remove "" freezeList))
  (setq thawList (vl-remove "" thawList))

  ;; Build an upper-cased freeze set for conflict detection in thaw pass.
  (setq freezeSet (mapcar 'strcase freezeList))

  (foreach layName freezeList
    (_freeze-one f dwg layName freezeSet)
  )

  (foreach layName thawList
    (_thaw-one f dwg layName freezeSet)
  )

  (command "_.QSAVE")
  (setq saveAfter (getvar "DBMOD"))
  (setq saveStatus (if (= saveAfter 0) "SAVE_OK" (strcat "SAVE_FAILED_DBMOD=" (itoa saveAfter))))

  (_log f dwg "<DWG_SAVE>" "SAVE" "T" "?" "?" "?" "?" "QSAVE" "?" saveStatus)

  (close f)
  (princ)
)

(defun c:BatchUpdateFiles (/)
  (c:BatchUpdate)
  (command "_.QUIT" "_N")
  (princ)
)
(princ)
"@
Set-Content -Path $updateLsp -Value $updateLspContent -Encoding ASCII

$updateScr = Join-Path $ToolDir "update.scr"
$updateScrLines = @(
  "FILEDIA",
  "0",
  "CMDDIA",
  "0",
  "PROXYNOTICE",
  "0",
  "SECURELOAD",
  "0",
  "(load `"$updateLspForLisp`")",
  "(c:BatchUpdateFiles)"
)
Set-Content -Path $updateScr -Value ($updateScrLines -join "`r`n") -Encoding ASCII

# ---------------- 7) EXECUTE UPDATES ----------------
$logFile = Join-Path ([Environment]::GetFolderPath("Desktop")) "LayerUpdateLog.txt"
"Starting Update at $(Get-Date)" | Out-File $logFile
"Files:" | Out-File $logFile -Append
$files | ForEach-Object { $_ | Out-File $logFile -Append }
"Layers to freeze: $($layersToFreeze -join ', ')" | Out-File $logFile -Append
"Layers to thaw:   $($layersToThaw -join ', ')" | Out-File $logFile -Append

Write-Host "PROGRESS: Updating $($files.Count) file(s)..."

$outLog = Join-Path $ToolDir "update_batch.out.txt"
$errLog = Join-Path $ToolDir "update_batch.err.txt"
$fileIndex = 0

foreach ($dwgFile in $files) {
  $fileIndex++
  Write-Host "PROGRESS: Processing $fileIndex of $($files.Count): $([IO.Path]::GetFileName($dwgFile))"
  if (Test-Path $outLog) { Remove-Item $outLog -Force }
  if (Test-Path $errLog) { Remove-Item $errLog -Force }

  $r = Invoke-AcadCore -DwgPath $dwgFile -ScriptPath $updateScr -OutLog $outLog -ErrLog $errLog -TimeoutSeconds $ProcessTimeoutSeconds

  if ($r.TimedOut) {
    Write-Host "    -> TIMEOUT. Killed." -ForegroundColor Red
    "FAILED (Timeout): $dwgFile" | Out-File $logFile -Append
  }
  elseif ($r.ExitCode -eq 0) {
    Write-Host "    -> ExitCode 0" -ForegroundColor Green
    "DONE (0): $dwgFile" | Out-File $logFile -Append
  }
  else {
    Write-Host "    -> ExitCode $($r.ExitCode)" -ForegroundColor Yellow
    "DONE ($($r.ExitCode)): $dwgFile" | Out-File $logFile -Append
  }
}

if ($files.Count -gt 0) {
  $outputFolder = Split-Path -Parent $files[0]
  if (-not [string]::IsNullOrWhiteSpace($outputFolder)) {
    Write-Host "PROGRESS: OUTPUT_FOLDER: $outputFolder"
  }
}

Write-Host "Done."
Write-Host "Report: $updateReport"
Write-Host "Logs: $ToolDir"
Write-Host "Summary log: $logFile"
