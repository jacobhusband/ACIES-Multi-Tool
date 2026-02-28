param(
  [string]$AcadCore,
  [string]$AutoDetectPaperSize = "true",
  [int]$ShrinkPercent = 100,
  [string]$FilesListPath = ""
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

function Ensure-WinFormsAssemblies {
  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing
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

function Convert-ToLispPath {
  param([string]$PathValue)
  if ([string]::IsNullOrWhiteSpace($PathValue)) { return "" }
  return ($PathValue -replace '\\', '/')
}

function Convert-ToLispQuotedList {
  param([string[]]$Values)
  $cleanValues = @(
    $Values |
      Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
      Select-Object -Unique
  )
  if ($cleanValues.Count -eq 0) {
    return '"ctextapp.arx"'
  }
  return ($cleanValues | ForEach-Object { '"' + ($_ -replace '"', '\"') + '"' }) -join " "
}

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
  $argsList = @("-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-File", $PSCommandPath)
  if ($AcadCore) { $argsList += @("-AcadCore", $AcadCore) }
  if ($PSBoundParameters.ContainsKey('AutoDetectPaperSize')) {
    $argsList += @("-AutoDetectPaperSize", $AutoDetectPaperSize)
  }
  if ($PSBoundParameters.ContainsKey('ShrinkPercent')) {
    $argsList += @("-ShrinkPercent", $ShrinkPercent)
  }
  if ($PSBoundParameters.ContainsKey('FilesListPath') -and -not [string]::IsNullOrWhiteSpace($FilesListPath)) {
    $argsList += @("-FilesListPath", $FilesListPath)
  }
  $child = Start-Process -FilePath $ps -ArgumentList $argsList -Wait -PassThru
  exit $child.ExitCode
}
# Validate that accoreconsole.exe exists
if ([string]::IsNullOrEmpty($acadCore) -or -not (Test-Path $acadCore)) {
  Write-Host "PROGRESS: ERROR: AutoCAD Core Console not found for versions 2020-2025."
  Write-Host "Please ensure AutoCAD is installed in the default 'C:\Program Files\Autodesk' directory."
  exit 1
}

$arcAlignedTextSupportCandidates = [System.Collections.Generic.List[string]]::new()
$acadInstallDir = Split-Path -Path $acadCore -Parent
if (-not [string]::IsNullOrWhiteSpace($acadInstallDir)) {
  $candidateFromExpress = Join-Path $acadInstallDir "Express\ctextapp.arx"
  $candidateFromRoot = Join-Path $acadInstallDir "ctextapp.arx"
  foreach ($candidate in @($candidateFromExpress, $candidateFromRoot)) {
    if (-not [string]::IsNullOrWhiteSpace($candidate)) {
      [void]$arcAlignedTextSupportCandidates.Add($candidate)
    }
  }
}
[void]$arcAlignedTextSupportCandidates.Add("ctextapp.arx")
$arcAlignedTextSupportCandidates = @(
  $arcAlignedTextSupportCandidates |
    Where-Object { $_ -and $_.Trim() } |
    Select-Object -Unique
)
$arcAlignedTextSupportCandidatesForLisp = @(
  $arcAlignedTextSupportCandidates | ForEach-Object { Convert-ToLispPath $_ }
)
$lispArcAlignedTextSupportCandidates = Convert-ToLispQuotedList $arcAlignedTextSupportCandidatesForLisp
$arcAlignedTextFailureMarker = "ACIES_ERROR: ARCALIGNEDTEXT_NOT_SUPPORTED"
$arcCheckPresentMarker = "ACIES_ARC_CHECK:PRESENT"
$arcCheckAbsentMarker = "ACIES_ARC_CHECK:ABSENT"
$arcModuleLoadSuccessMarker = "ACIES_ARC_MODULE_LOAD:SUCCESS"
$arcModuleLoadFailedMarker = "ACIES_ARC_MODULE_LOAD:FAILED"
$arcModuleLoadSkippedMarker = "ACIES_ARC_MODULE_LOAD:SKIPPED"
$plotDecisionContinueMarker = "ACIES_PLOT_DECISION:CONTINUE"
$plotDecisionSkipMarker = "ACIES_PLOT_DECISION:SKIP"
$preflightErrorMarker = "ACIES_PREFLIGHT_ERROR"
Write-Host "PROGRESS: ARCALIGNEDTEXT module candidates: $($arcAlignedTextSupportCandidates -join '; ')"

Ensure-WinFormsAssemblies

# --- Resolve DWG files: prefer preselected list, otherwise prompt ---
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
}

# Normalize to a string array so single-file runs still behave like multi-file runs.
$files = @($files)

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
$form.StartPosition = "Manual"
$form.TopMost = $true
$form.ShowInTaskbar = $true
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
Move-FormToPrimaryScreen $form

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

$form.add_Shown({
    $form.TopMost = $true
    $form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
    Move-FormToPrimaryScreen $form
    $form.Activate()
    $form.BringToFront()
    $comboBox.Focus()
  })

$form.add_Resize({
    if ($form.WindowState -eq [System.Windows.Forms.FormWindowState]::Minimized) {
      $form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
      Move-FormToPrimaryScreen $form
      $form.Activate()
      $form.BringToFront()
    }
  })

Write-Host "PROGRESS: Waiting for paper size confirmation..."
Write-Host "PROGRESS: Paper size dialog should be visible on the primary display."
Write-Host "PROGRESS: TRACE branch=paper_size_dialog"
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
"ARCALIGNEDTEXT module candidates: $($arcAlignedTextSupportCandidates -join '; ')" | Out-File $logFile -Append
"Output Folder: $batchOutputDir" | Out-File $logFile -Append
"Processing $($files.Count) files..." | Out-File $logFile -Append

# Initialize lists to track progress and generated files
$allGeneratedPdfs = [System.Collections.ArrayList]::new()
$failed = @()
$arcAlignedTextSupportFailures = @()
$noPdfOutputFailures = @()
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

(defun AciesLog (msg)
  (princ (strcat "\n" msg))
)

(defun AciesSafeArxLoad (module / res msg)
  (if (and module (> (strlen module) 0))
    (progn
      (setq res (vl-catch-all-apply 'arxload (list module)))
      (if (vl-catch-all-error-p res)
        (progn
          (setq msg (strcase (vl-catch-all-error-message res)))
          (or (wcmatch msg "*ALREADY*LOADED*")
              (wcmatch msg "*DUPLICATE*LOAD*"))
        )
        T
      )
    )
    nil
  )
)

(defun EnsureArcAlignedTextSupport (/ candidates candidate resolved loaded)
  (setq loaded nil)
  (setq candidates (list $lispArcAlignedTextSupportCandidates))
  (foreach candidate candidates
    (if (not loaded)
      (progn
        (setq resolved (findfile candidate))
        (if (not resolved)
          (setq resolved candidate)
        )
        (if (AciesSafeArxLoad resolved)
          (setq loaded T)
        )
      )
    )
  )
  loaded
)

(defun SafeRegenAll (/ result)
  (setq result (vl-catch-all-apply 'command (list "._REGENALL")))
  (if (vl-catch-all-error-p result)
    (AciesLog (strcat "REGENALL skipped: " (vl-catch-all-error-message result)))
  )
)

(defun EnsurePublishPreflight (/ hasArcAlignedText arcSupportLoaded arcSelectResult)
  (setvar "BACKGROUNDPLOT" 0)
  (setvar "FILEDIA" 0)
  (setvar "DEMANDLOAD" 3)
  (setvar "PROXYSHOW" 1)
  (SafeRefreshLinkedOLEs)
  (setq hasArcAlignedText nil)
  (setq arcSelectResult (vl-catch-all-apply 'ssget (list "_X" (list (cons 0 "ARCALIGNEDTEXT")))))
  (if (vl-catch-all-error-p arcSelectResult)
    (progn
      (AciesLog (strcat "${preflightErrorMarker}: " (vl-catch-all-error-message arcSelectResult)))
      (AciesLog "$plotDecisionSkipMarker")
      nil
    )
    (progn
      (setq hasArcAlignedText (and arcSelectResult (> (sslength arcSelectResult) 0)))
      (if hasArcAlignedText
        (progn
          (AciesLog "$arcCheckPresentMarker")
          (AciesLog "Loading ARCALIGNEDTEXT support module (ctextapp.arx)...")
          (setq arcSupportLoaded (EnsureArcAlignedTextSupport))
          (SafeRegenAll)
          (if arcSupportLoaded
            (progn
              (AciesLog "$arcModuleLoadSuccessMarker")
              (AciesLog "$plotDecisionContinueMarker")
              T
            )
            (progn
              (AciesLog "$arcModuleLoadFailedMarker")
              (AciesLog "$arcAlignedTextFailureMarker")
              (AciesLog "$plotDecisionSkipMarker")
              nil
            )
          )
        )
        (progn
          (AciesLog "$arcCheckAbsentMarker")
          (AciesLog "$arcModuleLoadSkippedMarker")
          (AciesLog "$plotDecisionContinueMarker")
          T
        )
      )
    )
  )
)

(defun c:PlotAllLayouts (/ main-dict layout-dict item layout-name pdfName preflightResult)
  (setq preflightResult (vl-catch-all-apply 'EnsurePublishPreflight '()))
  (if (vl-catch-all-error-p preflightResult)
    (progn
      (AciesLog (strcat "${preflightErrorMarker}: " (vl-catch-all-error-message preflightResult)))
      (AciesLog "$plotDecisionSkipMarker")
      (command "QUIT" "N")
    )
    (if preflightResult
      (progn
        (setq main-dict (namedobjdict))
        (setq layout-dict (dictsearch main-dict "ACAD_LAYOUT"))
        (foreach item layout-dict
          (if (= (car item) 3)
            (progn
              (setq layout-name (cdr item))
              (if (/= (strcase layout-name) "MODEL")
                (progn
                  (setvar "CTAB" layout-name)
                  (setq pdfName (strcat "$lispOutputDir\\" "$dwgNameWithoutExt" "-" layout-name ".pdf"))
                  (command "-PLOT" "Y" "" "DWG to PDF.pc3" "$selectedPaperSize" "I" "L" "N" "L" "1:1" "0.00,0.00" "Y" "510-monochrome.ctb" "Y" "N" "N" "N" pdfName "N" "Y")
                )
              )
            )
          )
        )
        (command "QUIT" "N")
      )
      (progn
        (AciesLog "Skipping plot because ARCALIGNEDTEXT support is unavailable in headless mode.")
        (command "QUIT" "N")
      )
    )
  )
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

  $existingPerDwgPdfs = @(Get-ChildItem -Path $batchOutputDir -Filter "$($dwgNameWithoutExt)-*.pdf" -ErrorAction SilentlyContinue)
  $existingPerDwgPdfSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($existingPdf in $existingPerDwgPdfs) {
    [void]$existingPerDwgPdfSet.Add($existingPdf.FullName)
  }

  $plotOutput = & $acadCore /i "$dwgPath" /s "$scriptFile" 2>&1 | Tee-Object -FilePath $logFile -Append
  $code = $LASTEXITCODE
  $plotOutputText = ($plotOutput | ForEach-Object { "$_" }) -join [Environment]::NewLine
  $plotOutputTextNormalized = $plotOutputText -replace ([char]0), ''
  $normalizedOutputLines = @(
    $plotOutput | ForEach-Object { ("$_" -replace ([char]0), '') }
  )
  $arcCheckPresent = $plotOutputTextNormalized -match [regex]::Escape($arcCheckPresentMarker)
  $arcCheckAbsent = $plotOutputTextNormalized -match [regex]::Escape($arcCheckAbsentMarker)
  $arcModuleLoadSuccess = $plotOutputTextNormalized -match [regex]::Escape($arcModuleLoadSuccessMarker)
  $arcModuleLoadFailed = $plotOutputTextNormalized -match [regex]::Escape($arcModuleLoadFailedMarker)
  $arcModuleLoadSkipped = $plotOutputTextNormalized -match [regex]::Escape($arcModuleLoadSkippedMarker)
  $plotDecisionContinue = $plotOutputTextNormalized -match [regex]::Escape($plotDecisionContinueMarker)
  $plotDecisionSkip = $plotOutputTextNormalized -match [regex]::Escape($plotDecisionSkipMarker)
  $missingArcAlignedTextSupport = $plotOutputTextNormalized -match [regex]::Escape($arcAlignedTextFailureMarker)
  $preflightError = $plotOutputTextNormalized -match [regex]::Escape($preflightErrorMarker)
  $preflightErrorLine = @(
    $normalizedOutputLines |
      Where-Object { $_ -like "*$preflightErrorMarker*" } |
      Select-Object -First 1
  )

  $arcCheckStatus = "unknown"
  if ($arcCheckPresent) { $arcCheckStatus = "present" }
  elseif ($arcCheckAbsent) { $arcCheckStatus = "absent" }

  $arcModuleLoadStatus = "unknown"
  if ($arcModuleLoadSuccess) { $arcModuleLoadStatus = "success" }
  elseif ($arcModuleLoadFailed) { $arcModuleLoadStatus = "failed" }
  elseif ($arcModuleLoadSkipped) { $arcModuleLoadStatus = "skipped" }

  $plotDecision = "unknown"
  if ($plotDecisionContinue) { $plotDecision = "continue" }
  elseif ($plotDecisionSkip) { $plotDecision = "skip" }

  $currentPerDwgPdfs = @(Get-ChildItem -Path $batchOutputDir -Filter "$($dwgNameWithoutExt)-*.pdf" -ErrorAction SilentlyContinue | Sort-Object FullName)
  $newPerDwgPdfs = @($currentPerDwgPdfs | Where-Object { -not $existingPerDwgPdfSet.Contains($_.FullName) })
  $newPerDwgPdfCount = $newPerDwgPdfs.Count

  "ARC_CHECK: $arcCheckStatus" | Out-File $logFile -Append
  "ARC_MODULE_LOAD: $arcModuleLoadStatus" | Out-File $logFile -Append
  "PLOT_DECISION: $plotDecision" | Out-File $logFile -Append
  "PDF_COUNT_AFTER_DWG: $newPerDwgPdfCount" | Out-File $logFile -Append
  Write-Host "TELEMETRY: $($dwgItem.Name): ARC_CHECK=$arcCheckStatus; ARC_MODULE_LOAD=$arcModuleLoadStatus; PLOT_DECISION=$plotDecision; PDF_COUNT_AFTER_DWG=$newPerDwgPdfCount"

  $preflightSkipped = ($plotDecision -eq "skip")
  $noPdfOutputFailure = (($plotDecision -eq "continue" -or $plotDecision -eq "unknown") -and $newPerDwgPdfCount -eq 0)

  if ($missingArcAlignedTextSupport -or ($preflightSkipped -and $arcCheckStatus -eq "present" -and $arcModuleLoadStatus -eq "failed")) {
    if (-not ($arcAlignedTextSupportFailures -contains $dwgPath)) {
      $arcAlignedTextSupportFailures += $dwgPath
    }
    Write-Host "PROGRESS: ERROR: $($dwgItem.Name) - ARCALIGNEDTEXT support missing (ctextapp.arx is not loadable by accoreconsole)."
    "ARCALIGNEDTEXT support failure detected in headless plotting output." | Out-File $logFile -Append
    "Hint: Ensure Express Tools is installed and ctextapp.arx is loadable by accoreconsole." | Out-File $logFile -Append
  }
  elseif ($preflightSkipped) {
    Write-Host "PROGRESS: ERROR: $($dwgItem.Name) - Plot skipped by preflight checks."
    "Plot skipped by preflight checks." | Out-File $logFile -Append
  }
  if ($preflightError) {
    $details = if ($preflightErrorLine) { $preflightErrorLine } else { "Preflight reported an unspecified AutoLISP error." }
    Write-Host "PROGRESS: ERROR: $($dwgItem.Name) - $details"
    "Preflight error detail: $details" | Out-File $logFile -Append
  }

  if ($noPdfOutputFailure) {
    if (-not ($noPdfOutputFailures -contains $dwgPath)) {
      $noPdfOutputFailures += $dwgPath
    }
    Write-Host "PROGRESS: ERROR: $($dwgItem.Name) - No PDFs were generated from layouts."
    "No PDFs were generated from layouts for this DWG." | Out-File $logFile -Append
  }

  if ($code -ne 0) {
    Write-Host "PROGRESS: ERROR: $($dwgItem.Name) - accoreconsole exited with code $code."
  }
  "ExitCode: $code" | Out-File $logFile -Append

  if ($code -ne 0 -or $missingArcAlignedTextSupport -or $preflightSkipped -or $preflightError -or $noPdfOutputFailure) {
    if (-not ($failed -contains $dwgPath)) {
      $failed += $dwgPath
    }
  }
  else {
    if ($newPerDwgPdfs.Count -gt 0) {
      $pdfPaths = @($newPerDwgPdfs | ForEach-Object { $_.FullName } | Sort-Object)
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
if ($arcAlignedTextSupportFailures.Count) {
  $arcFails = @($arcAlignedTextSupportFailures | Select-Object -Unique)
  $arcFailMsg = "ARCALIGNEDTEXT entities could not be plotted in headless mode because ctextapp.arx could not be loaded: $($arcFails -join ', ')"
  Write-Host "PROGRESS: ERROR: $arcFailMsg"
  Write-Host "PROGRESS: ERROR: Install/repair AutoCAD Express Tools (ctextapp.arx) for the selected AutoCAD Core Console."
  $arcFailMsg | Out-File $logFile -Append
  "Install/repair AutoCAD Express Tools and verify ctextapp.arx can be loaded by accoreconsole." | Out-File $logFile -Append
}
if ($noPdfOutputFailures.Count) {
  $noPdfFails = @($noPdfOutputFailures | Select-Object -Unique)
  $noPdfFailMsg = "One or more files reported plot success but produced zero layout PDFs: $($noPdfFails -join ', ')"
  Write-Host "PROGRESS: ERROR: $noPdfFailMsg"
  $noPdfFailMsg | Out-File $logFile -Append
}
if ($failed.Count) {
  $failedUnique = @($failed | Select-Object -Unique)
  $failMsg = "One or more files failed to plot: $($failedUnique -join ', ')"
  Write-Host "PROGRESS: ERROR: $failMsg"
  $failMsg | Out-File $logFile -Append
}

# Open the final output folder as the very last step
Invoke-Item $batchOutputDir

if ($failed.Count) {
  exit 1
}
