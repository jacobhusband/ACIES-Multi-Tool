param(
  [string]$AcadCore,
  [string]$DisciplineShort,
  [string]$FilesListPath = "",
  [string]$DefaultDirectory = "",
  [int]$StripXrefs = 1,
  [int]$SetByLayer = 1,
  [int]$Purge = 1,
  [int]$Audit = 1,
  [int]$HatchColor = 1,
  [int]$SkipAcad = 0
)

# removeXrefPaths.ps1
# Process DWGs -> accoreconsole loads your .NET DLL via NETLOAD -> runs STRIPREFPATHS
# All feedback is sent to the console for the main application to display.

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
      break # Stop searching once the latest version is found
    }
  }
}

$scriptRoot = $PSScriptRoot
$dll = Join-Path $scriptRoot "StripRefPaths.dll"

# Validation
if ([string]::IsNullOrEmpty($acadCore) -or -not (Test-Path $acadCore)) {
  Write-Host "PROGRESS: ERROR: AutoCAD Core Console not found for versions 2020-2025."
  Write-Host "Please ensure AutoCAD is installed in the default 'C:\Program Files\Autodesk' directory."
  exit 1
}

if ($StripXrefs -and -not (Test-Path $dll)) {
  Write-Host "PROGRESS: ERROR: Required component 'StripRefPaths.dll' not found at $dll"
  exit 1
}

# Relaunch in STA mode to guarantee the custom file picker can be sized/moved.
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
  $ps = (Get-Process -Id $PID).Path
  $argsList = @("-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-File", "`"$PSCommandPath`"")
  if ($AcadCore) { $argsList += @("-AcadCore", "`"$AcadCore`"") }
  if ($DisciplineShort) { $argsList += @("-DisciplineShort", "`"$DisciplineShort`"") }
  if ($FilesListPath) { $argsList += @("-FilesListPath", "`"$FilesListPath`"") }
  if ($DefaultDirectory) { $argsList += @("-DefaultDirectory", "`"$DefaultDirectory`"") }
  if ($PSBoundParameters.ContainsKey('StripXrefs')) { $argsList += @("-StripXrefs", $StripXrefs) }
  if ($PSBoundParameters.ContainsKey('SetByLayer')) { $argsList += @("-SetByLayer", $SetByLayer) }
  if ($PSBoundParameters.ContainsKey('Purge')) { $argsList += @("-Purge", $Purge) }
  if ($PSBoundParameters.ContainsKey('Audit')) { $argsList += @("-Audit", $Audit) }
  if ($PSBoundParameters.ContainsKey('HatchColor')) { $argsList += @("-HatchColor", $HatchColor) }
  if ($PSBoundParameters.ContainsKey('SkipAcad')) { $argsList += @("-SkipAcad", $SkipAcad) }
  $child = Start-Process -FilePath $ps -ArgumentList $argsList -Wait -PassThru
  exit $child.ExitCode
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Get-DisciplineShort([string]$value) {
  $short = ($value -split '[,;/\s]')[0]
  if ([string]::IsNullOrWhiteSpace($short)) { return "E" }
  return $short.Trim().ToUpper()
}

function Join-PathParts([string[]]$parts) {
  if (-not $parts -or $parts.Count -eq 0) { return "" }
  # Filter out empty strings (from UNC path splitting)
  $nonEmpty = $parts | Where-Object { $_ }
  if ($nonEmpty.Count -eq 0) { return "" }

  # Check if original path was UNC (starts with \\)
  # If first two parts were empty, it was a UNC path
  $isUnc = ($parts.Count -ge 2 -and $parts[0] -eq "" -and $parts[1] -eq "")

  if ($isUnc) {
    # Rebuild UNC path: \\server\share\...
    return "\\" + ($nonEmpty -join "\")
  }
  else {
    # Regular path
    $path = $nonEmpty[0]
    for ($i = 1; $i -lt $nonEmpty.Count; $i++) {
      $path = Join-Path -Path $path -ChildPath $nonEmpty[$i]
    }
    return $path
  }
}

function Find-FolderIndex([string[]]$parts, [string]$name) {
  for ($i = 0; $i -lt $parts.Count; $i++) {
    if ($parts[$i] -ieq $name) { return $i }
  }
  return -1
}

function Get-SheetIdFromName([string]$name) {
  # Prefer sheet IDs that start with A (ex: A00-00), then fall back to
  # single-letter dash patterns (ex: E-1-02-100).
  $preferredPattern = '(?i)\bA\d{1,3}-\d{1,3}[A-Z]?\b'
  $preferredMatch = [regex]::Match($name, $preferredPattern)
  if ($preferredMatch.Success) { return $preferredMatch.Value.ToUpper() }

  $fallbackPattern = '(?i)\b[A-Z]-\d{1,3}-\d{2}-\d+\b'
  $fallbackMatch = [regex]::Match($name, $fallbackPattern)
  if ($fallbackMatch.Success) { return $fallbackMatch.Value.ToUpper() }

  return $null
}

function Get-CleanFileName([string]$name) {
  $invalid = [IO.Path]::GetInvalidFileNameChars()
  $clean = $name
  foreach ($char in $invalid) {
    $clean = $clean -replace [regex]::Escape($char), ''
  }
  $clean = $clean.Trim()
  if ([string]::IsNullOrWhiteSpace($clean)) { return "Sheet" }
  return $clean
}

function Expand-ZipToTempFolder {
  param([Parameter(Mandatory = $true)][string]$ZipPath)

  $fullZipPath = [IO.Path]::GetFullPath($ZipPath)
  if (-not (Test-Path -LiteralPath $fullZipPath)) {
    throw "ZIP file not found: $fullZipPath"
  }

  $tempFolder = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath ("removeXrefPaths_zip_{0}" -f ([guid]::NewGuid().ToString("N")))
  New-Item -Path $tempFolder -ItemType Directory -Force | Out-Null

  try {
    Expand-Archive -LiteralPath $fullZipPath -DestinationPath $tempFolder -Force -ErrorAction Stop
  }
  catch {
    try {
      Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
      [System.IO.Compression.ZipFile]::ExtractToDirectory($fullZipPath, $tempFolder)
    }
    catch {
      Remove-TempExtractionFolderSafely -FolderPath $tempFolder
      throw "Unable to extract ZIP '$fullZipPath'. $($_.Exception.Message)"
    }
  }

  return $tempFolder
}

function Get-DwgFilesUnderFolder {
  param([string]$RootPath)

  if ([string]::IsNullOrWhiteSpace($RootPath) -or -not (Test-Path -LiteralPath $RootPath)) {
    return @()
  }

  return @(
    Get-ChildItem -LiteralPath $RootPath -Recurse -File -Filter "*.dwg" -ErrorAction SilentlyContinue |
      ForEach-Object { $_.FullName }
  )
}

function Remove-TempExtractionFolderSafely {
  param([string]$FolderPath)

  if ([string]::IsNullOrWhiteSpace($FolderPath)) { return }

  try {
    if (Test-Path -LiteralPath $FolderPath) {
      Remove-Item -LiteralPath $FolderPath -Recurse -Force -ErrorAction Stop
      Write-Host "PROGRESS: Removed temporary extraction folder: $FolderPath"
    }
  }
  catch {
    Write-Host "PROGRESS: WARNING: Failed to remove temporary extraction folder '$FolderPath'. $($_.Exception.Message)"
  }
}

function Show-SourceFileSelectionPrompt {
  param([string]$InitialDirectory = "")

  $promptForm = New-Object System.Windows.Forms.Form
  $promptForm.Text = "Select ZIP or DWG source file(s)"
  $promptForm.StartPosition = "CenterScreen"
  $promptForm.Size = New-Object System.Drawing.Size(600, 210)
  $promptForm.MinimumSize = New-Object System.Drawing.Size(600, 210)
  $promptForm.MaximumSize = New-Object System.Drawing.Size(600, 210)
  $promptForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
  $promptForm.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
  $promptForm.MaximizeBox = $false
  $promptForm.MinimizeBox = $false
  $promptForm.TopMost = $true
  $promptForm.ShowInTaskbar = $true

  $lblPrompt = New-Object System.Windows.Forms.Label
  $lblPrompt.Text = "Choose exactly one ZIP file or one or more DWG files. Mixed ZIP + DWG selection is not allowed."
  $lblPrompt.Location = New-Object System.Drawing.Point(16, 16)
  $lblPrompt.Size = New-Object System.Drawing.Size(560, 52)
  $lblPrompt.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
  $promptForm.Controls.Add($lblPrompt)

  $btnSelectFiles = New-Object System.Windows.Forms.Button
  $btnSelectFiles.Text = "Select ZIP or DWG Files..."
  $btnSelectFiles.Size = New-Object System.Drawing.Size(250, 44)
  $btnSelectFiles.Location = New-Object System.Drawing.Point(222, 106)
  $btnSelectFiles.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
  $btnSelectFiles.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
  $btnSelectFiles.FlatStyle = [System.Windows.Forms.FlatStyle]::System
  $promptForm.Controls.Add($btnSelectFiles)

  $btnExitPrompt = New-Object System.Windows.Forms.Button
  $btnExitPrompt.Text = "Exit"
  $btnExitPrompt.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
  $btnExitPrompt.Size = New-Object System.Drawing.Size(110, 44)
  $btnExitPrompt.Location = New-Object System.Drawing.Point(474, 106)
  $btnExitPrompt.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
  $btnExitPrompt.Font = New-Object System.Drawing.Font("Segoe UI", 10)
  $btnExitPrompt.FlatStyle = [System.Windows.Forms.FlatStyle]::System
  $promptForm.Controls.Add($btnExitPrompt)

  $promptForm.CancelButton = $btnExitPrompt
  $promptForm.AcceptButton = $btnSelectFiles

  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Title = "Select ZIP or DWG file(s)"
  $dlg.Filter = "CAD inputs (*.zip;*.dwg)|*.zip;*.dwg|ZIP files (*.zip)|*.zip|DWG files (*.dwg)|*.dwg"
  $dlg.Multiselect = $true
  $dlg.CheckFileExists = $true
  $dlg.CheckPathExists = $true
  $dlg.RestoreDirectory = $true
  if ($InitialDirectory -and (Test-Path $InitialDirectory)) {
    $dlg.InitialDirectory = $InitialDirectory
  }

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

  $promptResult = $promptForm.ShowDialog()
  if ($promptResult -eq [System.Windows.Forms.DialogResult]::OK) {
    $selectedFiles = @($promptForm.Tag)
    if ($selectedFiles.Count -gt 0) {
      return $selectedFiles
    }
  }
  return $null
}

function Show-DwgFileSelectionPrompt {
  param([string]$InitialDirectory = "")

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
  if ($InitialDirectory -and (Test-Path $InitialDirectory)) {
    $dlg.InitialDirectory = $InitialDirectory
  }

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

  $promptResult = $promptForm.ShowDialog()
  if ($promptResult -eq [System.Windows.Forms.DialogResult]::OK) {
    $selectedFiles = @($promptForm.Tag)
    if ($selectedFiles.Count -gt 0) {
      return $selectedFiles
    }
  }
  return $null
}

Write-Host "PROGRESS: Staging cleanup commands..."
# Stage AutoLISP CLEANUP command so each drawing is tidied after ref paths are stripped
$lisp = Join-Path $env:TEMP "cleanup_tmp.lsp"
$setByLayerBlock = ""
if ($SetByLayer) {
  $setByLayerBlock = @'
  ;; Use SETBYLAYER on everything in the current space
  ;;   "Y" => Change ByBlock to ByLayer?
  ;;   "Y" => Include blocks?
  (command
    "_.-SETBYLAYER" "All" "" "Y" "Y")

'@
}

$hatchBlock = ""
if ($HatchColor) {
  $hatchBlock = @'
  ;; Change color of non-nested modelspace hatches to 9 (excluding WALL layers)
  ;; This runs AFTER SETBYLAYER so the color override is preserved
  ;; ssget "X" with filter: HATCH entities in modelspace (410 = "Model")
  (setq ss (ssget "X" '((0 . "HATCH") (410 . "Model"))))
  (if ss
    (progn
      (setq i 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i))
        (setq entData (entget ent))
        (setq layerName (cdr (assoc 8 entData)))
        ;; Check if layer name does NOT contain "WALL" (case-insensitive)
        (if (not (wcmatch (strcase layerName) "*WALL*"))
          (progn
            ;; Change color to 9 by modifying entity data
            (if (assoc 62 entData)
              ;; Color property exists, replace it
              (setq entData (subst (cons 62 9) (assoc 62 entData) entData))
              ;; Color property doesn't exist, add it
              (setq entData (append entData (list (cons 62 9))))
            )
            (entmod entData)
          )
        )
        (setq i (1+ i))
      )
    )
  )

'@
}

$purgeBlock = ""
if ($Purge) {
  $purgeBlock = @'
  ;; Purge the entire drawing; "All", "*" (everything), "No" to confirm each
  (command "_.-PURGE" "All" "*" "N")

'@
}

$auditBlock = ""
if ($Audit) {
  $auditBlock = @'
  ;; Audit the drawing; "Yes" to fix errors
  (command "_.AUDIT" "Y")

'@
}

$lispContent = @"
(defun c:CLEANUP (/ oldCMDECHO oldTab ss i ent entData layerName)
  ;; Remember current command echo setting, then turn it off:
  (setq oldCMDECHO (getvar "CMDECHO")
        oldTab     (getvar "CTAB"))
  (setvar "CMDECHO" 0)

  ;; Ensure commands run from Model space
  (if (/= oldTab "Model")
    (command "_.MODEL")
  )

$setByLayerBlock$hatchBlock$purgeBlock$auditBlock
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
"@
Set-Content -Encoding ASCII -Path $lisp -Value $lispContent

# Make .scr that runs the requested steps
$script = Join-Path $env:TEMP "run_STRIP.scr"
$lispPathForScript = ($lisp -replace '\\', '/')
$scriptLines = @(
  "CMDECHO 1",
  "FILEDIA 0",
  "SECURELOAD 0"
)
if ($StripXrefs) {
  $scriptLines += "NETLOAD"
  $scriptLines += "`"$dll`""
  $scriptLines += "STRIPREFPATHS"
}

$runCleanup = $SetByLayer -or $Purge -or $Audit -or $HatchColor
if ($runCleanup) {
  $scriptLines += "(load `"$lispPathForScript`")"
  $scriptLines += "CLEANUP"
}

$scriptLines += "QSAVE"
$scriptLines += "QUIT"
$scriptContent = $scriptLines -join "`r`n"
Set-Content -Encoding ASCII -Path $script -Value $scriptContent

# Resolve files: prefer preselected list, otherwise open custom picker.
$files = @()
$zipSelected = $false
$zipOutputRoot = ""
$tempExtractionFolders = @()
$cancelled = $false

try {
  if (-not [string]::IsNullOrWhiteSpace($FilesListPath) -and (Test-Path $FilesListPath)) {
    $files = @(
      Get-Content -Path $FilesListPath -Encoding UTF8 |
        Where-Object { $_ -and $_.Trim() -and (Test-Path $_.Trim()) } |
        ForEach-Object { [IO.Path]::GetFullPath($_.Trim()) }
    )
    if ($files.Count -gt 0) {
      Write-Host "PROGRESS: Using $($files.Count) DWG file(s) from auto-selected project folder."
    }
    else {
      Write-Host "PROGRESS: Provided files list was empty. Opening file picker..."
    }
  }
  elseif (-not [string]::IsNullOrWhiteSpace($FilesListPath)) {
    Write-Host "PROGRESS: Provided files list path not found. Opening file picker..."
  }

  if (-not $files -or $files.Count -eq 0) {
    while ($true) {
      $sourceSelections = Show-SourceFileSelectionPrompt -InitialDirectory $DefaultDirectory
      if (-not $sourceSelections -or @($sourceSelections).Count -eq 0) {
        Write-Host "PROGRESS: Cancelled. No files selected."
        $cancelled = $true
        break
      }

      $resolvedSelections = @(
        @($sourceSelections) |
          ForEach-Object { $_.ToString().Trim() } |
          Where-Object { $_ -and (Test-Path -LiteralPath $_) } |
          ForEach-Object { [IO.Path]::GetFullPath($_) }
      )
      if ($resolvedSelections.Count -eq 0) {
        Write-Host "PROGRESS: No valid files selected. Choose one ZIP or DWG file(s)."
        continue
      }

      $zipSelections = @(
        $resolvedSelections | Where-Object { [IO.Path]::GetExtension($_) -ieq ".zip" }
      )
      $dwgSelections = @(
        $resolvedSelections | Where-Object { [IO.Path]::GetExtension($_) -ieq ".dwg" }
      )
      $unsupportedSelectionCount = $resolvedSelections.Count - ($zipSelections.Count + $dwgSelections.Count)
      if ($unsupportedSelectionCount -gt 0) {
        Write-Host "PROGRESS: ERROR: Only ZIP and DWG files are supported. Select one ZIP file or DWG files only."
        continue
      }

      if ($zipSelections.Count -gt 0 -and $dwgSelections.Count -gt 0) {
        Write-Host "PROGRESS: ERROR: Mixed selection is not allowed. Select exactly one ZIP file or one or more DWG files."
        continue
      }

      if ($zipSelections.Count -gt 1) {
        Write-Host "PROGRESS: ERROR: Select only one ZIP file at a time, or select DWG files directly."
        continue
      }

      if ($zipSelections.Count -eq 1) {
        $zipPath = $zipSelections[0]
        Write-Host "PROGRESS: ZIP source selected. Extracting archive..."

        try {
          $zipExtractRoot = Expand-ZipToTempFolder -ZipPath $zipPath
          $tempExtractionFolders += $zipExtractRoot
          Write-Host "PROGRESS: Extracted ZIP to temporary folder: $zipExtractRoot"
        }
        catch {
          Write-Host "PROGRESS: ERROR: Failed to extract ZIP file. $($_.Exception.Message)"
          continue
        }

        $zipDwgs = Get-DwgFilesUnderFolder -RootPath $zipExtractRoot
        if ($zipDwgs.Count -eq 0) {
          Write-Host "PROGRESS: No DWG files were found inside the selected ZIP. Choose another ZIP or select DWGs directly."
          Remove-TempExtractionFolderSafely -FolderPath $zipExtractRoot
          continue
        }

        Write-Host "PROGRESS: Found $($zipDwgs.Count) DWG file(s) in extracted ZIP. Select DWG file(s) to process."
        $selectedFromZip = Show-DwgFileSelectionPrompt -InitialDirectory $zipExtractRoot
        if (-not $selectedFromZip -or @($selectedFromZip).Count -eq 0) {
          Write-Host "PROGRESS: Cancelled. No files selected."
          $cancelled = $true
          break
        }

        $zipRootPrefix = [IO.Path]::GetFullPath($zipExtractRoot).TrimEnd('\') + '\'
        $selectedZipDwgs = @(
          @($selectedFromZip) |
            ForEach-Object { $_.ToString().Trim() } |
            Where-Object {
              $_ -and
              (Test-Path -LiteralPath $_) -and
              ([IO.Path]::GetExtension($_) -ieq ".dwg")
            } |
            ForEach-Object { [IO.Path]::GetFullPath($_) } |
            Where-Object { $_.StartsWith($zipRootPrefix, [System.StringComparison]::OrdinalIgnoreCase) }
        )

        if ($selectedZipDwgs.Count -eq 0) {
          Write-Host "PROGRESS: No DWG files from the extracted ZIP were selected. Choose source files again."
          Remove-TempExtractionFolderSafely -FolderPath $zipExtractRoot
          continue
        }
        if ($selectedZipDwgs.Count -ne @($selectedFromZip).Count) {
          Write-Host "PROGRESS: Some selected files were outside the extracted ZIP folder and were ignored."
        }

        $zipSelected = $true
        $zipBaseName = [IO.Path]::GetFileNameWithoutExtension($zipPath)
        $zipParentFolder = Split-Path -Parent $zipPath
        $zipOutputRoot = Join-Path -Path $zipParentFolder -ChildPath ("{0}_Prepared" -f $zipBaseName)
        if (-not (Test-Path -LiteralPath $zipOutputRoot)) {
          New-Item -Path $zipOutputRoot -ItemType Directory -Force | Out-Null
          Write-Host "PROGRESS: Created ZIP output folder at $zipOutputRoot"
        }
        Write-Host "PROGRESS: ZIP output folder: $zipOutputRoot"

        $files = @($selectedZipDwgs)
        break
      }

      if ($dwgSelections.Count -gt 0) {
        $files = @($dwgSelections)
        break
      }

      Write-Host "PROGRESS: ERROR: Select exactly one ZIP file or one or more DWG files."
    }
  }

  if ($cancelled) {
    return
  }

  if (-not $files -or $files.Count -eq 0) {
    Write-Host "PROGRESS: Cancelled. No files selected."
    return
  }

  # Normalize to string array for single-file selection consistency.
  $files = @($files)
  Write-Host "PROGRESS: Processing $($files.Count) file(s)..."

  $disciplineShort = Get-DisciplineShort $DisciplineShort
  $failed = @()
  $i = 0
  foreach ($dwg in $files) {
    $i++
    $name = [IO.Path]::GetFileName($dwg)
    Write-Host "PROGRESS: Preparing $i of $($files.Count): $name"

    try {
      $fullPath = [IO.Path]::GetFullPath($dwg)
      $parts = $fullPath -split '[\\/]'

      $sourcePath = $fullPath
      $workingPath = $fullPath
      $targetDir = Split-Path -Parent $fullPath
      $copiedFromArch = $false

      if ($zipSelected) {
        $targetDir = $zipOutputRoot
        if (-not (Test-Path -LiteralPath $targetDir)) {
          New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
        }
      }
      else {
        $xrefIndex = Find-FolderIndex $parts 'Xrefs'
        if ($xrefIndex -lt 0) { $xrefIndex = Find-FolderIndex $parts 'Xref' }
        $archIndex = Find-FolderIndex $parts 'Arch'

        if ($xrefIndex -ge 0) {
          # Already in Xrefs/Xref folder - use as-is
          $workingPath = $fullPath
        }
        elseif ($archIndex -ge 0) {
          # In Arch folder - copy to Xrefs folder (directly, no subfolders)
          if ($archIndex -le 0) {
            $projectRoot = $parts[0]
          }
          else {
            $projectRoot = Join-PathParts $parts[0..($archIndex - 1)]
          }
          # Put files directly in Xrefs folder (no subfolder preservation)
          $targetDir = Join-Path -Path $projectRoot -ChildPath "Xrefs"
          if (-not (Test-Path $targetDir)) {
            # Only create if it doesn't exist
            New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
            Write-Host "PROGRESS: Created Xrefs folder at $targetDir"
          }
          $copiedFromArch = $true
          Write-Host "PROGRESS: Preparing Arch source for Xrefs transfer: $name"
        }
      }

      $baseName = [IO.Path]::GetFileNameWithoutExtension($sourcePath)
      $sheetId = Get-SheetIdFromName $baseName
      if ($sheetId) {
        $targetBase = Get-CleanFileName $sheetId
      }
      else {
        $targetBase = Get-CleanFileName $baseName
      }
      $targetName = "{0} ({1})" -f $targetBase, $disciplineShort
      $targetPath = Join-Path -Path $targetDir -ChildPath "$targetName.dwg"

      if ($copiedFromArch) {
        $incomingName = "__incoming__{0}.dwg" -f ([guid]::NewGuid().ToString('N'))
        $incomingPath = Join-Path -Path $targetDir -ChildPath $incomingName
        Copy-Item -LiteralPath $sourcePath -Destination $incomingPath -Force
        $workingPath = $incomingPath
        Write-Host "PROGRESS: Staged Arch source in Xrefs as $incomingName"
        Write-Host "PROGRESS: Final target in Xrefs: $([IO.Path]::GetFileName($targetPath))"
      }

      $existing = Get-ChildItem -LiteralPath $targetDir -File |
        Where-Object { $_.BaseName -ieq $targetName }
      Write-Host "PROGRESS: Checking exact target-name collisions for '$targetName' in $targetDir"
      foreach ($item in $existing) {
        if ($item.FullName -ne $workingPath) {
          # Archive the existing file instead of deleting it
          $archiveDir = Join-Path -Path $targetDir -ChildPath "Archive"
          if (-not (Test-Path $archiveDir)) {
            New-Item -Path $archiveDir -ItemType Directory -Force | Out-Null
            Write-Host "PROGRESS: Created Archive folder at $archiveDir"
          }
          $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
          $archiveName = "{0}_{1}{2}" -f $item.BaseName, $timestamp, $item.Extension
          $archivePath = Join-Path -Path $archiveDir -ChildPath $archiveName
          Move-Item -LiteralPath $item.FullName -Destination $archivePath -Force
          Write-Host "PROGRESS: Archived existing file to $archiveName"
        }
      }

      if ($workingPath -ne $targetPath) {
        Move-Item -LiteralPath $workingPath -Destination $targetPath -Force
        $workingPath = $targetPath
      }
    }
    catch {
      Write-Host "PROGRESS: ERROR: Failed to prepare $name - $_"
      $failed += $dwg
      continue
    }

    Write-Host "PROGRESS: Processing $i of $($files.Count): $([IO.Path]::GetFileName($workingPath))"

    if ($SkipAcad) {
      Write-Host "PROGRESS: SkipAcad enabled. Skipping accoreconsole execution for $([IO.Path]::GetFileName($workingPath))."
      continue
    }

    # IMPORTANT: no /ld here; we NETLOAD via the .scr
    # The 2>&1 redirects stderr to stdout, so the Python wrapper can catch errors.
    & $acadCore /i "$workingPath" /s "$script" 2>&1
    $code = $LASTEXITCODE
    if ($code -ne 0) { $failed += $workingPath }
  }

  Write-Progress -Activity "Processing DWGs" -Completed

  if ($failed.Count) {
    Write-Host "PROGRESS: ERROR: $($failed.Count) of $($files.Count) file(s) failed to process."
  }
  else {
    Write-Host "PROGRESS: Successfully processed $($files.Count) drawing(s)."
  }
}
finally {
  $uniqueTempExtractionFolders = @(
    $tempExtractionFolders |
      Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
      Sort-Object -Unique
  )
  foreach ($tempFolder in $uniqueTempExtractionFolders) {
    Remove-TempExtractionFolderSafely -FolderPath $tempFolder
  }
}
