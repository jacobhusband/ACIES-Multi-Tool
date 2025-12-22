param(
  [string]$AcadCore
)

Write-Host "PROGRESS: Initializing script..."

# --- AUTOCAD VERSION DETECTION ---
if ($AcadCore -and (Test-Path -Path $AcadCore)) {
  $acadCore = $AcadCore
  Write-Host "PROGRESS: Using specified AutoCAD Core Console: $AcadCore"
}
else {
  $acadCore = $null
  $years = 2025, 2024, 2023, 2022, 2021, 2020

  foreach ($year in $years) {
    $possiblePath = "C:\Program Files\Autodesk\AutoCAD $year\accoreconsole.exe"
    if (Test-Path -Path $possiblePath) {
      $acadCore = $possiblePath
      Write-Host "PROGRESS: Found AutoCAD $year Core Console."
      break
    }
  }
}

# Relaunch in STA for file picker
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
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

# --- Let the user select DWGs via a file explorer dialog ---
Write-Host "PROGRESS: Waiting for file selection..."
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$dlg = New-Object System.Windows.Forms.OpenFileDialog
$dlg.Title = "Select DWG file(s) to update layer freeze states"
$dlg.Filter = "DWG files (*.dwg)|*.dwg|All files (*.*)|*.*"
$dlg.Multiselect = $true
$dlg.InitialDirectory = [Environment]::GetFolderPath("Desktop")
if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK -or -not $dlg.FileNames) {
  Write-Host "PROGRESS: ERROR: No files selected."; exit 1
}
$files = $dlg.FileNames

function Test-DwgLocked {
  param([string]$Path)
  $dir = Split-Path -Path $Path -Parent
  $name = [IO.Path]::GetFileNameWithoutExtension($Path)
  $dwl = Join-Path $dir "$name.dwl"
  $dwl2 = Join-Path $dir "$name.dwl2"
  if (Test-Path $dwl -or Test-Path $dwl2) {
    return $true
  }
  try {
    $stream = [System.IO.File]::Open(
      $Path,
      [System.IO.FileMode]::Open,
      [System.IO.FileAccess]::ReadWrite,
      [System.IO.FileShare]::None
    )
    $stream.Close()
    return $false
  }
  catch {
    return $true
  }
}

$lockedFiles = @()
$readOnlyFiles = @()
foreach ($file in $files) {
  $item = Get-Item -LiteralPath $file -ErrorAction SilentlyContinue
  if (-not $item) {
    $lockedFiles += $file
    continue
  }
  if ($item.IsReadOnly) {
    $readOnlyFiles += $file
    continue
  }
  if (Test-DwgLocked $file) {
    $lockedFiles += $file
  }
}

if ($readOnlyFiles.Count -gt 0 -or $lockedFiles.Count -gt 0) {
  if ($readOnlyFiles.Count -gt 0) {
    Write-Host "PROGRESS: ERROR: These files are read-only and cannot be saved:"
    $readOnlyFiles | ForEach-Object { Write-Host "PROGRESS: - $_" }
  }
  if ($lockedFiles.Count -gt 0) {
    Write-Host "PROGRESS: ERROR: These files appear to be open or locked:"
    $lockedFiles | ForEach-Object { Write-Host "PROGRESS: - $_" }
  }
  Write-Host "PROGRESS: ERROR: Please close the drawings and retry."
  exit 1
}

# --- Extract Layers from ALL selected files ---
$tempExtractDir = Join-Path $env:TEMP "acad_layer_extract"
if (Test-Path $tempExtractDir) { Remove-Item $tempExtractDir -Recurse -Force }
New-Item -ItemType Directory -Path $tempExtractDir | Out-Null

$extractLisp = Join-Path $tempExtractDir "extract_layers.lsp"
$extractLispContent = @"
(vl-load-com)
(setvar "CMDECHO" 0)
(setvar "FILEDIA" 0)
(defun c:ExtractLayers (/ lay flags frozen)
  (princ "\nDEBUG_LISP_START\n")
  (setq lay (tblnext "LAYER" T))
  (while lay
    (setq flags (cdr (assoc 70 lay)))
    (if (null flags) (setq flags 0))
    (setq frozen (if (= (logand flags 1) 1) "FROZEN" "THAWED"))
    (princ (strcat "\nLAYER_FOUND:" (cdr (assoc 2 lay)) "|" frozen))
    (setq lay (tblnext "LAYER"))
  )
  (command "_QUIT" "Y")
)
"@
Set-Content -Encoding ASCII -Path $extractLisp -Value $extractLispContent

$extractScr = Join-Path $tempExtractDir "extract_layers.scr"
$extractScrContent = "(load `"$($extractLisp -replace '\\', '/')`")`nExtractLayers`n"
Set-Content -Encoding ASCII -Path $extractScr -Value $extractScrContent

$fileIndex = 0
$extractionOutputs = @()
foreach ($file in $files) {
  $fileIndex++
  Write-Host "PROGRESS: Scanning layers ($fileIndex of $($files.Count)): $(Split-Path $file -Leaf)"

  $pInfo = New-Object System.Diagnostics.ProcessStartInfo
  $pInfo.FileName = $acadCore
  $pInfo.Arguments = "/i `"$file`" /s `"$extractScr`""
  $pInfo.RedirectStandardOutput = $true
  $pInfo.RedirectStandardError = $true
  $pInfo.StandardOutputEncoding = [System.Text.Encoding]::Unicode
  $pInfo.StandardErrorEncoding = [System.Text.Encoding]::Unicode
  $pInfo.UseShellExecute = $false
  $pInfo.CreateNoWindow = $true

  $p = New-Object System.Diagnostics.Process
  $p.StartInfo = $pInfo
  [void]$p.Start()

  $output = $p.StandardOutput.ReadToEnd()
  $err = $p.StandardError.ReadToEnd()
  $p.WaitForExit()

  $cleanOutput = $output -replace '\0', ''
  $cleanErr = $err -replace '\0', ''
  $extractionOutputs += $cleanOutput
  $extractionOutputs += $cleanErr
}

# Aggregate unique layers + frozen state
$allLayers = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$frozenLayers = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$thawedLayers = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($output in $extractionOutputs) {
  if ($output) {
    $lines = $output -split "\r?\n"
    foreach ($line in $lines) {
      if ($line -match "LAYER_FOUND:(.+)") {
        $payload = $matches[1].Trim()
        $parts = $payload -split "\|", 2
        $layerName = $parts[0].Trim()
        if (-not [string]::IsNullOrWhiteSpace($layerName)) {
          $state = if ($parts.Count -gt 1) { $parts[1].Trim().ToUpperInvariant() } else { "THAWED" }
          [void]$allLayers.Add($layerName)
          if ($state -eq "FROZEN") {
            [void]$frozenLayers.Add($layerName)
          }
          else {
            [void]$thawedLayers.Add($layerName)
          }
        }
      }
    }
  }
}

$defaultFreezeLayers = [System.Collections.Generic.List[string]]::new()
$defaultKeepLayers = [System.Collections.Generic.List[string]]::new()
foreach ($lay in $allLayers) {
  if ($frozenLayers.Contains($lay) -and -not $thawedLayers.Contains($lay)) {
    [void]$defaultFreezeLayers.Add($lay)
  }
  else {
    [void]$defaultKeepLayers.Add($lay)
  }
}

if ($allLayers.Count -eq 0) {
  Write-Host "PROGRESS: WARNING: No layers were extracted. The list will be empty."
}

$allKeepLayers = [System.Collections.Generic.List[string]]::new()
foreach ($lay in $defaultKeepLayers) {
  if (-not [string]::IsNullOrWhiteSpace($lay)) {
    [void]$allKeepLayers.Add([string]$lay)
  }
}
$allFreezeLayers = [System.Collections.Generic.List[string]]::new()
foreach ($lay in $defaultFreezeLayers) {
  if (-not [string]::IsNullOrWhiteSpace($lay)) {
    [void]$allFreezeLayers.Add([string]$lay)
  }
}

# --- Let the user select layers to Freeze/Thaw ---
Write-Host "PROGRESS: Waiting for user input..."
$form = New-Object System.Windows.Forms.Form
$form.Text = "Layer Freeze Settings"
$form.Size = New-Object System.Drawing.Size(600, 500)
$form.StartPosition = "CenterScreen"

# List Labels
$lblKeep = New-Object System.Windows.Forms.Label
$lblKeep.Location = New-Object System.Drawing.Point(10, 45)
$lblKeep.Text = "Layers to KEEP (Unfrozen)"
$lblKeep.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($lblKeep)

$lblFreeze = New-Object System.Windows.Forms.Label
$lblFreeze.Location = New-Object System.Drawing.Point(320, 45)
$lblFreeze.Text = "Layers to FREEZE"
$lblFreeze.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($lblFreeze)

# Filter box
$lblLayerFilter = New-Object System.Windows.Forms.Label
$lblLayerFilter.Location = New-Object System.Drawing.Point(10, 10)
$lblLayerFilter.Text = "Filter layers:"
$lblLayerFilter.Size = New-Object System.Drawing.Size(80, 20)
$form.Controls.Add($lblLayerFilter)

$txtLayerFilter = New-Object System.Windows.Forms.TextBox
$txtLayerFilter.Location = New-Object System.Drawing.Point(95, 8)
$txtLayerFilter.Size = New-Object System.Drawing.Size(475, 20)
$form.Controls.Add($txtLayerFilter)

# ListBoxes
$listKeep = New-Object System.Windows.Forms.ListBox
$listKeep.Location = New-Object System.Drawing.Point(10, 70)
$listKeep.Size = New-Object System.Drawing.Size(250, 300)
$listKeep.SelectionMode = "MultiExtended"
$form.Controls.Add($listKeep)

$listFreeze = New-Object System.Windows.Forms.ListBox
$listFreeze.Location = New-Object System.Drawing.Point(320, 70)
$listFreeze.Size = New-Object System.Drawing.Size(250, 300)
$listFreeze.SelectionMode = "MultiExtended"
$form.Controls.Add($listFreeze)

function Remove-Layer {
  param(
    [System.Collections.Generic.List[string]]$list,
    [string]$item
  )
  for ($i = $list.Count - 1; $i -ge 0; $i--) {
    if ($list[$i] -ieq $item) {
      $list.RemoveAt($i)
      return
    }
  }
}

function Add-LayerUnique {
  param(
    [System.Collections.Generic.List[string]]$list,
    [string]$item
  )
  if ($list -notcontains $item) {
    [void]$list.Add($item)
  }
}

function Update-LayerList {
  param(
    [System.Windows.Forms.ListBox]$listBox,
    [System.Collections.Generic.List[string]]$source,
    [string]$filterText
  )
  $listBox.BeginUpdate()
  try {
    $listBox.Items.Clear()
    $items = $source
    if (-not [string]::IsNullOrWhiteSpace($filterText)) {
      $needle = $filterText.Trim()
      $items = $source | Where-Object {
        $_.IndexOf($needle, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
      }
    }
    foreach ($item in ($items | Sort-Object)) {
      [void]$listBox.Items.Add($item)
    }
  }
  finally {
    $listBox.EndUpdate()
  }
}

$refreshLayerLists = {
  Update-LayerList $listKeep $allKeepLayers $txtLayerFilter.Text
  Update-LayerList $listFreeze $allFreezeLayers $txtLayerFilter.Text
}

$txtLayerFilter.Add_TextChanged({ & $refreshLayerLists })
& $refreshLayerLists

$moveToFreeze = {
  $selected = @($listKeep.SelectedItems)
  foreach ($item in $selected) {
    Remove-Layer $allKeepLayers $item
    Add-LayerUnique $allFreezeLayers $item
  }
  & $refreshLayerLists
}

$moveToKeep = {
  $selected = @($listFreeze.SelectedItems)
  foreach ($item in $selected) {
    Remove-Layer $allFreezeLayers $item
    Add-LayerUnique $allKeepLayers $item
  }
  & $refreshLayerLists
}

# Buttons to move
$btnToFreeze = New-Object System.Windows.Forms.Button
$btnToFreeze.Text = ">"
$btnToFreeze.Location = New-Object System.Drawing.Point(275, 170)
$btnToFreeze.Size = New-Object System.Drawing.Size(35, 30)
$btnToFreeze.Add_Click({ & $moveToFreeze })
$form.Controls.Add($btnToFreeze)

$btnToKeep = New-Object System.Windows.Forms.Button
$btnToKeep.Text = "<"
$btnToKeep.Location = New-Object System.Drawing.Point(275, 210)
$btnToKeep.Size = New-Object System.Drawing.Size(35, 30)
$btnToKeep.Add_Click({ & $moveToKeep })
$form.Controls.Add($btnToKeep)

# Double-click handlers
$listKeep.Add_MouseDoubleClick({ & $moveToFreeze })
$listFreeze.Add_MouseDoubleClick({ & $moveToKeep })

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(260, 390)
$okButton.Size = New-Object System.Drawing.Size(80, 30)
$okButton.Text = "OK"
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$form.Topmost = $true
$result = $form.ShowDialog()

if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
  Write-Host "PROGRESS: ERROR: Operation cancelled by user."; exit 1
}

$layersToFreeze = @($allFreezeLayers | Sort-Object)
$layersToKeep = @($allKeepLayers | Sort-Object)

# --- Create SCR file with direct commands (more reliable than LISP) ---
$scriptFile = Join-Path $env:TEMP "freeze_layers.scr"
$scriptLines = @()

# Add CMDECHO off
$scriptLines += "CMDECHO"
$scriptLines += "0"
$scriptLines += "FILEDIA"
$scriptLines += "0"

# Find a safe layer to set as current
$safeLayer = $null
foreach ($lay in $layersToKeep) {
  $safeLayer = $lay
  break
}
if (-not $safeLayer) {
  $safeLayer = "0"
}

# Set safe layer as current
$scriptLines += "-LAYER"
$scriptLines += "S"
$scriptLines += $safeLayer
$scriptLines += ""

# Thaw all layers that should be kept
foreach ($lay in $layersToKeep) {
  $scriptLines += "-LAYER"
  $scriptLines += "T"
  $scriptLines += $lay
  $scriptLines += ""
}

# Freeze all layers that should be frozen
foreach ($lay in $layersToFreeze) {
  $scriptLines += "-LAYER"
  $scriptLines += "F"
  $scriptLines += $lay
  $scriptLines += ""
}

# Save and quit
$scriptLines += "QSAVE"
$scriptLines += ""
$scriptLines += "QUIT"
$scriptLines += "Y"

$scriptContent = $scriptLines -join "`n"
Set-Content -Encoding ASCII -Path $scriptFile -Value $scriptContent

# --- PROCESSING SETUP ---
Write-Host "PROGRESS: Preparing to update $($files.Count) file(s)..."
$baseLogDir = Join-Path -Path ([Environment]::GetFolderPath("MyDocuments")) -ChildPath "AutoCAD Layer Updates"
$timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$runOutputDir = Join-Path -Path $baseLogDir -ChildPath $timestamp
New-Item -ItemType Directory -Force -Path $runOutputDir | Out-Null

$logFile = Join-Path $runOutputDir "_LayerFreezeLog.txt"
"===== Layer Freeze Started: $(Get-Date -f 'yyyy-MM-dd HH:mm:ss') =====" | Out-File $logFile
"AutoCAD Core Used: $acadCore" | Out-File $logFile -Append
"Safe Layer (Current): $safeLayer" | Out-File $logFile -Append
"Layers to Freeze ($($layersToFreeze.Count)): $($layersToFreeze -join ', ')" | Out-File $logFile -Append
"Layers to Keep ($($layersToKeep.Count)): $($layersToKeep -join ', ')" | Out-File $logFile -Append
"Processing $($files.Count) files..." | Out-File $logFile -Append
"" | Out-File $logFile -Append
"Script content:" | Out-File $logFile -Append
$scriptContent | Out-File $logFile -Append
"" | Out-File $logFile -Append

$failed = @()
$i = 0
foreach ($dwgPath in $files) {
  $i++
  $dwgItem = Get-Item $dwgPath
  Write-Host "PROGRESS: Updating layers ($i of $($files.Count)): $($dwgItem.Name)"

  "===== $(Get-Date -f 'yyyy-MM-dd HH:mm:ss') Start: $($dwgItem.Name) =====" | Out-File $logFile -Append
  & $acadCore /i "$dwgPath" /s "$scriptFile" 2>&1 | Tee-Object -FilePath $logFile -Append
  $code = $LASTEXITCODE
  "ExitCode: $code" | Out-File $logFile -Append

  if ($code -ne 0) {
    $failed += $dwgPath
  }
  "===== $(Get-Date -f 'yyyy-MM-dd HH:mm:ss') Done: $($dwgItem.Name) =====" | Out-File $logFile -Append
}

"===== Layer Freeze Finished: $(Get-Date -f 'yyyy-MM-dd HH:mm:ss') =====" | Out-File $logFile -Append

if ($failed.Count) {
  $failMsg = "One or more files failed to update: $($failed -join ', ')"
  Write-Host "PROGRESS: ERROR: $failMsg"
  $failMsg | Out-File $logFile -Append
}
else {
  Write-Host "PROGRESS: Successfully updated $($files.Count) drawing(s)."
}

Write-Host "PROGRESS: Log saved to $logFile"