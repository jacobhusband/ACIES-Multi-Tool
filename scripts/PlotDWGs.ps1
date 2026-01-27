param(
  [string]$AcadCore,
  [string]$AutoDetectPaperSize = "true",
  [int]$ShrinkPercent = 100
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

$AutoDetectPaperSize = Convert-ToBool $AutoDetectPaperSize $true

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
$detectSizeScriptPath = Join-Path $scriptRoot "detect_pdf_size.py"
$shrinkScriptPath = Join-Path $scriptRoot "shrink_pdf.py"
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
  $argsList = @("-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-File", "`"$PSCommandPath`"")
  if ($AcadCore) { $argsList += @("-AcadCore", "`"$AcadCore`"") }
  if ($PSBoundParameters.ContainsKey('AutoDetectPaperSize')) {
    $argsList += @("-AutoDetectPaperSize", $AutoDetectPaperSize)
  }
  if ($PSBoundParameters.ContainsKey('ShrinkPercent')) {
    $argsList += @("-ShrinkPercent", $ShrinkPercent)
  }
  Start-Process -FilePath $ps -ArgumentList $argsList -Wait
  exit
}
# Validate that accoreconsole.exe exists
if ([string]::IsNullOrEmpty($acadCore) -or -not (Test-Path $acadCore)) {
  Write-Host "PROGRESS: ERROR: AutoCAD Core Console not found for versions 2020-2025."
  Write-Host "Please ensure AutoCAD is installed in the default 'C:\Program Files\Autodesk' directory."
  exit 1
}

# --- Let the user select DWGs FIRST via a file explorer dialog ---
Write-Host "PROGRESS: Waiting for user input..."
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

# --- Detect paper size from existing project PDFs ---
$detectedPaperSize = ""
$detectionStatus = "not_checked"
if ($AutoDetectPaperSize) {
  if (Test-Path $detectSizeScriptPath) {
    Write-Host "PROGRESS: Detecting paper size from existing PDFs..."
    try {
      # Quote the path in case it contains spaces
      $dwgPathArg = "`"$($files[0])`""
      $detectOutput = & $pythonExecutable $detectSizeScriptPath $dwgPathArg 2>&1

      # Check if there was an error (stderr output will be error records)
      $errorOutput = $detectOutput | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] }
      $stdOutput = $detectOutput | Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] }

      if ($stdOutput) {
        $detectedPaperSize = ($stdOutput | Out-String).Trim()
      }

      if ($detectedPaperSize) {
        Write-Host "PROGRESS: Detected paper size: $detectedPaperSize"
        $detectionStatus = "detected"
      } else {
        if ($errorOutput) {
          Write-Host "PROGRESS: Detection error: $($errorOutput | Out-String)"
        }
        Write-Host "PROGRESS: No matching paper size found in project PDFs."
        $detectionStatus = "no_match"
      }
    }
    catch {
      Write-Host "PROGRESS: Could not detect paper size: $_"
      $detectionStatus = "error"
      $detectedPaperSize = ""
    }
  } else {
    Write-Host "PROGRESS: Detection script not found at: $detectSizeScriptPath"
    $detectionStatus = "script_missing"
  }
}
else {
  Write-Host "PROGRESS: Auto-detect disabled. Please enter paper size manually."
  $detectionStatus = "disabled"
}

# --- Let the user select/confirm a paper size ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Select Paper Size"
$form.Size = New-Object System.Drawing.Size(450, 200)
$form.StartPosition = "CenterScreen"

# Status label to show detection result
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(10, 15)
$statusLabel.Size = New-Object System.Drawing.Size(420, 20)
$statusLabel.Font = New-Object System.Drawing.Font($statusLabel.Font.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
switch ($detectionStatus) {
  "detected" {
    $statusLabel.Text = "Paper size detected from existing project PDFs"
    $statusLabel.ForeColor = [System.Drawing.Color]::Green
  }
  "no_match" {
    $statusLabel.Text = "No matching PDFs found in project folder"
    $statusLabel.ForeColor = [System.Drawing.Color]::DarkOrange
  }
  "disabled" {
    $statusLabel.Text = "Auto-detect disabled"
    $statusLabel.ForeColor = [System.Drawing.Color]::Gray
  }
  default {
    $statusLabel.Text = "Auto-detection not available"
    $statusLabel.ForeColor = [System.Drawing.Color]::Gray
  }
}
$form.Controls.Add($statusLabel)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10, 45)
$label.Size = New-Object System.Drawing.Size(420, 20)
if ($detectedPaperSize) {
  $label.Text = "Confirm the detected size or select a different one:"
} elseif ($AutoDetectPaperSize) {
  $label.Text = "Please select a paper size for plotting:"
} else {
  $label.Text = "Please enter or select a paper size for plotting:"
}
$form.Controls.Add($label)

$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.Location = New-Object System.Drawing.Point(10, 70)
$comboBox.Size = New-Object System.Drawing.Size(410, 20)
if ($AutoDetectPaperSize) {
  $comboBox.DropDownStyle = "DropDownList"
}
else {
  $comboBox.DropDownStyle = "DropDown"
}

# Add paper sizes with checkmark for detected size
$selectedIndex = 0
$index = 0
foreach ($size in $paperSizes) {
  if ($detectedPaperSize -and $size -eq $detectedPaperSize) {
    [void] $comboBox.Items.Add("$size  [Detected]")
    $selectedIndex = $index
  } else {
    [void] $comboBox.Items.Add($size)
  }
  $index++
}
if ($AutoDetectPaperSize -and $comboBox.Items.Count -gt 0) {
  $comboBox.SelectedIndex = $selectedIndex
}
$form.Controls.Add($comboBox)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(175, 110)
$okButton.Size = New-Object System.Drawing.Size(75, 23)
$okButton.Text = "OK"
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$form.Topmost = $true
$result = $form.ShowDialog()

if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
  Write-Host "PROGRESS: ERROR: Operation cancelled by user."; exit
}

# Remove the " [Detected]" suffix if present to get the actual paper size
$selectedPaperSize = ($comboBox.Text -replace '\s+\[Detected\]$', '').Trim()
if ([string]::IsNullOrWhiteSpace($selectedPaperSize)) {
  Write-Host "PROGRESS: ERROR: No paper size selected."
  exit 1
}

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

    $shrinkPercentInt = [int]$ShrinkPercent
    if ($shrinkPercentInt -lt 5) { $shrinkPercentInt = 5 }
    if ($shrinkPercentInt -gt 100) { $shrinkPercentInt = 100 }
    if ($shrinkPercentInt -lt 100) {
      if (Test-Path $shrinkScriptPath) {
        $shrunkName = "combined-shrunk-$shrinkPercentInt-percent.pdf"
        $shrunkPath = Join-Path $batchOutputDir $shrunkName
        Write-Host "PROGRESS: Shrinking combined PDF to $shrinkPercentInt%..."
        $shrinkOutput = & $pythonExecutable $shrinkScriptPath "`"$combinedPdfPath`"" "`"$shrunkPath`"" $shrinkPercentInt 2>&1
        if ($LASTEXITCODE -eq 0) {
          "Shrunk PDF created: $shrunkName" | Out-File $logFile -Append
          "$shrinkOutput" | Out-File $logFile -Append
        } else {
          Write-Host "PROGRESS: ERROR: PDF shrinking failed."
          "PDF shrinking failed. Output from Python script:" | Out-File $logFile -Append
          "$shrinkOutput" | Out-File $logFile -Append
        }
      } else {
        Write-Host "PROGRESS: ERROR: 'shrink_pdf.py' not found."
        "'shrink_pdf.py' not found; skipped shrinking." | Out-File $logFile -Append
      }
    }
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
