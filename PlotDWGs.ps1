param(
  [string]$AcadCore
)

# --- SCRIPT CONFIGURATION ---
# Logic to find AutoCAD Core Console
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
      break # Stop searching once the latest version is found
    }
  }
}

# Name for the combined PDF file
$combinedPdfName = "combined.pdf"
# Name of the Python executable (can be python, python3, or a full path)
$pythonExecutable = "python"

# --- DEFINE AVAILABLE PAPER SIZES ---
$paperSizes = @(
  "ARCH full bleed E1 (30.00 x 42.00 Inches)",
  "ARCH full bleed D (24.00 x 36.00 Inches)",
  "ANSI full bleed D (22.00 x 34.00 Inches)"
)

# --- SCRIPT AND PYTHON VALIDATION ---
# Get the directory where this script is located
$scriptRoot = $PSScriptRoot
$pythonScriptPath = Join-Path $scriptRoot "merge_pdfs.py"
# Check if the Python script exists
if (-not (Test-Path $pythonScriptPath)) {
  Write-Host "PROGRESS: ERROR: 'merge_pdfs.py' not found."
  exit 1
}
# Check if Python is available in the system's PATH
$pythonCheck = Get-Command $pythonExecutable -ErrorAction SilentlyContinue
if (-not $pythonCheck) {
  Write-Host "PROGRESS: ERROR: Python executable ('$pythonExecutable') not found in PATH."
  exit 1
}
# Relaunch in STA mode for the file picker dialog to work correctly
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
  $ps = (Get-Process -Id $PID).Path
  Start-Process -FilePath $ps -ArgumentList @("-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-File", "`"$PSCommandPath`"") -Wait
  exit
}
# Validate that accoreconsole.exe exists
if ([string]::IsNullOrEmpty($acadCore) -or -not (Test-Path $acadCore)) {
  Write-Host "PROGRESS: ERROR: AutoCAD Core Console not found for versions 2020-2025."
  Write-Host "Please ensure AutoCAD is installed in the default 'C:\Program Files\Autodesk' directory."
  exit 1
}

# --- Let the user select DWGs via a file explorer dialog ---
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$dlg = New-Object System.Windows.Forms.OpenFileDialog
$dlg.Title = "Select DWG file(s) to plot"
$dlg.Filter = "DWG files (*.dwg)|*.dwg|All files (*.*)|*.*"
$dlg.Multiselect = $true
$dlg.InitialDirectory = [Environment]::GetFolderPath("Desktop")
if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK -or -not $dlg.FileNames) {
  Write-Host "PROGRESS: ERROR: No files selected."; exit
}
$files = $dlg.FileNames

# --- Extract Layers from ALL selected files ---
$tempExtractDir = Join-Path $env:TEMP "acad_extract"
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
    
  # Run accoreconsole and capture output to variable
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

  # Strip nulls in case the console output is UTF-16 interpreted as ANSI
  $cleanOutput = $output -replace '\0', ''
  $cleanErr = $err -replace '\0', ''
  $extractionOutputs += $cleanOutput
  $extractionOutputs += $cleanErr
}

# Aggregate unique layers
# --- DEBUG: DUMP RAW OUTPUT ---
$debugDumpPath = Join-Path $env:TEMP "debug_raw_output.txt"
"--- RAW OUTPUT DUMP START ---" | Out-File $debugDumpPath
foreach ($o in $extractionOutputs) {
  "--- CHUNK ---" | Out-File $debugDumpPath -Append
  $o | Out-File $debugDumpPath -Append
}
"--- RAW OUTPUT DUMP END ---" | Out-File $debugDumpPath -Append
Write-Host "DEBUG: Raw output dumped to $debugDumpPath"

# Aggregate unique layers + frozen state
$keepLayers = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$frozenLayers = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($output in $extractionOutputs) {
  if ($output) {
    # Split by lines (handles LF or CRLF)
    $lines = $output -split "\r?\n"
    foreach ($line in $lines) {
      # Match our specific prefix
      if ($line -match "LAYER_FOUND:(.+)") {
        $payload = $matches[1].Trim()
        $parts = $payload -split "\|", 2
        $layerName = $parts[0].Trim()
        if (-not [string]::IsNullOrWhiteSpace($layerName)) {
          $state = if ($parts.Count -gt 1) { $parts[1].Trim().ToUpperInvariant() } else { "THAWED" }
          if ($state.StartsWith("FROZEN")) {
            [void]$frozenLayers.Add($layerName)
          }
          else {
            [void]$keepLayers.Add($layerName)
          }
        }
      }
    }
  }
}

# If a layer appears thawed anywhere, default it to Keep.
foreach ($lay in $keepLayers) {
  [void]$frozenLayers.Remove($lay)
}

$availableLayers = @($keepLayers | Sort-Object)
$availableFrozenLayers = @($frozenLayers | Sort-Object)

if ($availableLayers.Count -eq 0 -and $availableFrozenLayers.Count -eq 0) {
  Write-Host "PROGRESS: WARNING: No layers were extracted. The list will be empty."
}

$allKeepLayers = [System.Collections.Generic.List[string]]::new()
foreach ($lay in $availableLayers) {
  if (-not [string]::IsNullOrWhiteSpace($lay)) {
    [void]$allKeepLayers.Add([string]$lay)
  }
}
$allFreezeLayers = [System.Collections.Generic.List[string]]::new()
foreach ($lay in $availableFrozenLayers) {
  if (-not [string]::IsNullOrWhiteSpace($lay)) {
    [void]$allFreezeLayers.Add([string]$lay)
  }
}

# --- Let the user select settings (Paper Size + Layers to Freeze) ---
Write-Host "PROGRESS: Waiting for user input..."
$form = New-Object System.Windows.Forms.Form
$form.Text = "Plotting Settings"
$form.Size = New-Object System.Drawing.Size(600, 500)
$form.StartPosition = "CenterScreen"

# Paper Size
$labelSize = New-Object System.Windows.Forms.Label
$labelSize.Location = New-Object System.Drawing.Point(10, 10)
$labelSize.Size = New-Object System.Drawing.Size(200, 20)
$labelSize.Text = "Select Paper Size:"
$form.Controls.Add($labelSize)

$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.Location = New-Object System.Drawing.Point(10, 30)
$comboBox.Size = New-Object System.Drawing.Size(560, 20)
$comboBox.DropDownStyle = "DropDownList"
$paperSizes | ForEach-Object { [void] $comboBox.Items.Add($_) }
$comboBox.SelectedIndex = 0
$form.Controls.Add($comboBox)

# List Labels
$lblKeep = New-Object System.Windows.Forms.Label
$lblKeep.Location = New-Object System.Drawing.Point(10, 95)
$lblKeep.Text = "Layers to KEEP (Unfrozen)"
$lblKeep.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($lblKeep)

$lblFreeze = New-Object System.Windows.Forms.Label
$lblFreeze.Location = New-Object System.Drawing.Point(320, 95)
$lblFreeze.Text = "Layers to FREEZE"
$lblFreeze.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($lblFreeze)

# Filter box
$lblLayerFilter = New-Object System.Windows.Forms.Label
$lblLayerFilter.Location = New-Object System.Drawing.Point(10, 70)
$lblLayerFilter.Text = "Filter layers:"
$lblLayerFilter.Size = New-Object System.Drawing.Size(80, 20)
$form.Controls.Add($lblLayerFilter)

$txtLayerFilter = New-Object System.Windows.Forms.TextBox
$txtLayerFilter.Location = New-Object System.Drawing.Point(95, 68)
$txtLayerFilter.Size = New-Object System.Drawing.Size(475, 20)
$form.Controls.Add($txtLayerFilter)

# ListBoxes
$listKeep = New-Object System.Windows.Forms.ListBox
$listKeep.Location = New-Object System.Drawing.Point(10, 120)
$listKeep.Size = New-Object System.Drawing.Size(250, 270)
$listKeep.SelectionMode = "MultiExtended"
$form.Controls.Add($listKeep)

$listFreeze = New-Object System.Windows.Forms.ListBox
$listFreeze.Location = New-Object System.Drawing.Point(320, 120)
$listFreeze.Size = New-Object System.Drawing.Size(250, 270)
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
$btnToFreeze.Location = New-Object System.Drawing.Point(275, 210)
$btnToFreeze.Size = New-Object System.Drawing.Size(35, 30)
$btnToFreeze.Add_Click({ & $moveToFreeze })
$form.Controls.Add($btnToFreeze)

$btnToKeep = New-Object System.Windows.Forms.Button
$btnToKeep.Text = "<"
$btnToKeep.Location = New-Object System.Drawing.Point(275, 250)
$btnToKeep.Size = New-Object System.Drawing.Size(35, 30)
$btnToKeep.Add_Click({ & $moveToKeep })
$form.Controls.Add($btnToKeep)

# Double-click handlers
$listKeep.Add_MouseDoubleClick({
    & $moveToFreeze
  })
$listFreeze.Add_MouseDoubleClick({
    & $moveToKeep
  })

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(260, 410)
$okButton.Size = New-Object System.Drawing.Size(80, 30)
$okButton.Text = "OK"
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$form.Topmost = $true
$result = $form.ShowDialog()

if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
  Write-Host "PROGRESS: ERROR: Operation cancelled by user."; exit
}

$selectedPaperSize = $comboBox.SelectedItem
$layersToFreeze = @($allFreezeLayers | Sort-Object)

# --- BATCH PROCESSING SETUP ---
Write-Host "PROGRESS: Preparing to plot $($files.Count) file(s)..."
$basePlotDir = Join-Path -Path ([Environment]::GetFolderPath("MyDocuments")) -ChildPath "AutoCAD Plots"
$timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$firstFileParentDirName = (Get-Item -Path $files[0]).Directory.Parent.Name
$batchOutputDir = Join-Path -Path (Join-Path -Path $basePlotDir -ChildPath $firstFileParentDirName) -ChildPath $timestamp
New-Item -ItemType Directory -Force -Path $batchOutputDir | Out-Null

# --- SINGLE LOG FILE SETUP ---
$logFile = Join-Path $batchOutputDir "_BatchPlotLog.txt"
"===== Batch Plot Started: $(Get-Date -f 'yyyy-MM-dd HH:mm:ss') =====" | Out-File $logFile
"Selected Paper Size: $selectedPaperSize" | Out-File $logFile -Append
"AutoCAD Core Used: $acadCore" | Out-File $logFile -Append
"Output Folder: $batchOutputDir" | Out-File $logFile -Append
"Processing $($files.Count) files..." | Out-File $logFile -Append

# Initialize lists to track progress and generated files
$allGeneratedPdfs = [System.Collections.ArrayList]::new()
$failed = @()
$i = 0

# --- Main Processing Loop (Plotting ONLY) ---
foreach ($dwgPath in $files) {
  $i++
  $dwgItem = Get-Item $dwgPath
  $dwgNameWithoutExt = $dwgItem.BaseName
  Write-Host "PROGRESS: Plotting $i of $($files.Count): $($dwgItem.Name)"
    
  "===== $(Get-Date -f 'yyyy-MM-dd HH:mm:ss') Start Plotting: $($dwgItem.Name) =====" | Out-File $logFile -Append
    
  # Create AutoLISP file for plotting
  $lispFile = Join-Path $env:TEMP "plot_layouts.lsp"
  $lispOutputDir = $batchOutputDir -replace '\\', '\\'
  
  $lispFreezeLogic = ""
  if ($layersToFreeze.Count -gt 0) {
    $lispFreezeLogic = "(setq freezeList '("
    foreach ($lay in $layersToFreeze) {
      # Escape backslashes and double quotes for LISP string
      $escapedLay = $lay -replace '\\', '\\\\' -replace '"', '\"'
      $lispFreezeLogic += "`"$escapedLay`" "
    }
    $lispFreezeLogic += "))"
    # Safely freeze if layer exists
    $lispFreezeLogic += "`n  (foreach lay freezeList "
    $lispFreezeLogic += " (if (tblsearch `"LAYER`" lay)"
    $lispFreezeLogic += " (vl-catch-all-apply 'vla-put-Freeze (list (vla-item (vla-get-Layers (vla-get-ActiveDocument (vlax-get-acad-object))) lay) :vlax-true))"
    $lispFreezeLogic += " ))"
  }

  $lispContent = @"
(defun UpdateBlockOLELinks (blk / res)
  (vlax-for obj blk
    (if (equal (vla-get-ObjectName obj) "AcDbOle2Frame")
      (progn
        (setq res (vl-catch-all-apply 'vla-Update (list obj)))
        (if (vl-catch-all-error-p res)
          (princ (strcat "\nOLE update failed: " (vl-catch-all-error-message res)))
          (setq *ole-updated-count* (1+ *ole-updated-count*))
        )
      )
    )
  )
)

(defun RefreshLinkedOLEs (/ acad doc)
  (vl-load-com)
  (setq acad (vlax-get-acad-object))
  (setq doc (vla-get-ActiveDocument acad))
  (setq *ole-updated-count* 0)
  (UpdateBlockOLELinks (vla-get-ModelSpace doc))
  (vlax-for lay (vla-get-Layouts doc)
    (if (/= (strcase (vla-get-Name lay)) "MODEL")
      (UpdateBlockOLELinks (vla-get-Block lay))
    )
  )
  (princ (strcat "\nOLE links refreshed: " (itoa *ole-updated-count*)))
)

(defun SafeRefreshLinkedOLEs (/ res)
  (setq res (vl-catch-all-apply 'RefreshLinkedOLEs '()))
  (if (vl-catch-all-error-p res)
    (princ (strcat "\nOLE refresh skipped: " (vl-catch-all-error-message res)))
  )
)



(defun c:PlotAllLayouts ()
  (setvar "BACKGROUNDPLOT" 0)
  (setvar "FILEDIA" 0)
  (SafeRefreshLinkedOLEs)
  $lispFreezeLogic
  (setq main-dict (namedobjdict))
  (setq layout-dict (dictsearch main-dict "ACAD_LAYOUT"))
  (foreach item layout-dict
    (if (= (car item) 3)
      (progn
        (setq layout-name (cdr item))
        (if (/= (strcase layout-name) "MODEL")
          (progn
            (setvar 'CTAB layout-name)
            (setq pdfName (strcat "$lispOutputDir\\" "$dwgNameWithoutExt" "-" layout-name ".pdf"))
            (command "-PLOT" "Y" "" "DWG to PDF.pc3" "$selectedPaperSize" "I" "L" "N" "L" "1:1" "0.00,0.00" "Y" "510-monochrome.ctb" "Y" "N" "N" "N" pdfName "N" "Y")
          )
        )
      )
    )
  )
  (command "QUIT" "Y")
  (princ)
)
"@
  Set-Content -Encoding ASCII -Path $lispFile -Value $lispContent
    
  $scriptFile = Join-Path $env:TEMP "run_plot.scr"
  $lispPathForScript = $lispFile -replace '\\', '/'
  $scriptContent = @"
(load "$lispPathForScript")
PlotAllLayouts
"@
  Set-Content -Encoding ASCII -Path $scriptFile -Value $scriptContent
    
  & $acadCore /i "$dwgPath" /s "$scriptFile" 2>&1 | Tee-Object -FilePath $logFile -Append
  $code = $LASTEXITCODE
  "ExitCode: $code" | Out-File $logFile -Append
    
  if ($code -ne 0) {
    $failed += $dwgPath
  }
  else {
    $individualPdfs = @(Get-ChildItem -Path $batchOutputDir -Filter "$($dwgNameWithoutExt)-*.pdf")
    if ($individualPdfs.Count -gt 0) {
      $pdfPaths = @($individualPdfs | ForEach-Object { $_.FullName } | Sort-Object)
      $allGeneratedPdfs.AddRange($pdfPaths)
    }
  }
  "===== $(Get-Date -f 'yyyy-MM-dd HH:mm:ss') Done: $($dwgItem.Name) =====" | Out-File $logFile -Append
}

# --- FINAL MERGE (After the loop) ---
if ($allGeneratedPdfs.Count -gt 0) {
  Write-Host "PROGRESS: Combining $($allGeneratedPdfs.Count) generated PDFs..."
  $combinedPdfPath = Join-Path $batchOutputDir $combinedPdfName
   
  $pythonArgs = @(
    "`"$combinedPdfPath`""
  ) + $allGeneratedPdfs.ForEach({ "`"$_`"" })
    
  $mergeOutput = & $pythonExecutable $pythonScriptPath $pythonArgs 2>&1
   
  if ($LASTEXITCODE -eq 0) {
    "Python script executed successfully." | Out-File $logFile -Append
    "$mergeOutput" | Out-File $logFile -Append
  }
  else {
    Write-Host "PROGRESS: ERROR: PDF merging failed."
    "PDF Merging Failed. Output from Python script:" | Out-File $logFile -Append
    "$mergeOutput" | Out-File $logFile -Append
  }
}
else {
  Write-Host "PROGRESS: No PDFs were generated to combine."
  "No PDFs were generated to merge." | Out-File $logFile -Append
}

# --- CLEANUP & NOTIFICATION ---
"===== Batch Plot Finished: $(Get-Date -f 'yyyy-MM-dd HH:mm:ss') =====" | Out-File $logFile -Append
if ($failed.Count) {
  $failMsg = "One or more files failed to plot: $($failed -join ', ')"
  Write-Host "PROGRESS: ERROR: $failMsg"
  $failMsg | Out-File $logFile -Append
}

# Open the final output folder as the very last step
Invoke-Item $batchOutputDir
