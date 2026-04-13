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

# Relaunch in STA mode so Windows file dialogs work reliably.
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

function Ensure-ArchiveDirectory {
  param([Parameter(Mandatory = $true)][string]$TargetDir)

  $archiveDir = Join-Path -Path $TargetDir -ChildPath "Archive"
  if (-not (Test-Path -LiteralPath $archiveDir)) {
    New-Item -Path $archiveDir -ItemType Directory -Force | Out-Null
    Write-Host "PROGRESS: Created Archive folder at $archiveDir"
  }
  return $archiveDir
}

function Move-FileToArchive {
  param(
    [Parameter(Mandatory = $true)][string]$FilePath,
    [Parameter(Mandatory = $true)][string]$ArchiveRoot,
    [string]$Message = "Archived existing file to"
  )

  $item = Get-Item -LiteralPath $FilePath -ErrorAction Stop
  $archiveDir = Ensure-ArchiveDirectory -TargetDir $ArchiveRoot
  $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
  $archiveName = "{0}_{1}{2}" -f $item.BaseName, $timestamp, $item.Extension
  $archivePath = Join-Path -Path $archiveDir -ChildPath $archiveName
  Move-Item -LiteralPath $item.FullName -Destination $archivePath -Force
  Write-Host "PROGRESS: $Message $archiveName"
  return $archivePath
}

function New-IncomingWorkingCopy {
  param(
    [Parameter(Mandatory = $true)][string]$SourcePath,
    [Parameter(Mandatory = $true)][string]$TargetDir,
    [string]$SourceLabel = "source"
  )

  $incomingName = "__incoming__{0}{1}" -f ([guid]::NewGuid().ToString('N')), [IO.Path]::GetExtension($SourcePath)
  $incomingPath = Join-Path -Path $TargetDir -ChildPath $incomingName
  Copy-Item -LiteralPath $SourcePath -Destination $incomingPath -Force
  Write-Host "PROGRESS: Staged $SourceLabel in Xrefs as $incomingName"
  return $incomingPath
}

function Show-DwgFileDialog {
  param([string]$InitialDirectory = "")

  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Title = "Select DWG file(s)"
  $dlg.Filter = "DWG files (*.dwg)|*.dwg"
  $dlg.Multiselect = $true
  $dlg.CheckFileExists = $true
  $dlg.CheckPathExists = $true
  $dlg.RestoreDirectory = $true
  if ($InitialDirectory -and (Test-Path -LiteralPath $InitialDirectory)) {
    $dlg.InitialDirectory = $InitialDirectory
  }

  $result = $dlg.ShowDialog()
  if ($result -eq [System.Windows.Forms.DialogResult]::OK -and $dlg.FileNames.Count -gt 0) {
    return [string[]]$dlg.FileNames
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

# Resolve files: prefer preselected list, otherwise open the DWG picker.
$files = @()
if (-not [string]::IsNullOrWhiteSpace($FilesListPath) -and (Test-Path $FilesListPath)) {
  $files = @(
    Get-Content -Path $FilesListPath -Encoding UTF8 |
      Where-Object {
        $_ -and
        $_.Trim() -and
        (Test-Path -LiteralPath $_.Trim()) -and
        ([IO.Path]::GetExtension($_.Trim()) -ieq ".dwg")
      } |
      ForEach-Object { [IO.Path]::GetFullPath($_.Trim()) } |
      Select-Object -Unique
  )
  if ($files.Count -gt 0) {
    Write-Host "PROGRESS: Using $($files.Count) DWG file(s) from auto-selected project folder."
  }
  else {
    Write-Host "PROGRESS: Provided files list was empty. Opening DWG file picker..."
  }
}
elseif (-not [string]::IsNullOrWhiteSpace($FilesListPath)) {
  Write-Host "PROGRESS: Provided files list path not found. Opening DWG file picker..."
}

if (-not $files -or $files.Count -eq 0) {
  Write-Host "PROGRESS: Opening DWG file picker..."
  $sourceSelections = Show-DwgFileDialog -InitialDirectory $DefaultDirectory
  if (-not $sourceSelections -or @($sourceSelections).Count -eq 0) {
    Write-Host "PROGRESS: Cancelled. No files selected."
    return
  }

  $files = @(
    @($sourceSelections) |
      ForEach-Object { $_.ToString().Trim() } |
      Where-Object {
        $_ -and
        (Test-Path -LiteralPath $_) -and
        ([IO.Path]::GetExtension($_) -ieq ".dwg")
      } |
      ForEach-Object { [IO.Path]::GetFullPath($_) } |
      Select-Object -Unique
  )
}

if (-not $files -or $files.Count -eq 0) {
  Write-Host "PROGRESS: Cancelled. No DWG files selected."
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
    $stagedInXrefs = $false
    $archiveSelectedSource = $false
    $stageLabel = ""

    $xrefIndex = Find-FolderIndex $parts 'Xrefs'
    if ($xrefIndex -lt 0) { $xrefIndex = Find-FolderIndex $parts 'Xref' }
    $archIndex = Find-FolderIndex $parts 'Arch'

    if ($xrefIndex -ge 0) {
      $relativeXrefParts = @()
      if (($xrefIndex + 1) -le ($parts.Count - 2)) {
        $relativeXrefParts = @($parts[($xrefIndex + 1)..($parts.Count - 2)])
      }
      if (@($relativeXrefParts | Where-Object { $_ -ieq 'Archive' }).Count -gt 0) {
        throw "Selected DWG is already inside the Xrefs\Archive folder. Choose a DWG from Arch or the active Xrefs folder instead."
      }

      $targetDir = Join-PathParts $parts[0..$xrefIndex]
      $stagedInXrefs = $true
      $archiveSelectedSource = $true
      $stageLabel = "Xrefs source"
      Write-Host "PROGRESS: Preparing Xrefs source for Xrefs cleanup: $name"
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
      $stagedInXrefs = $true
      $stageLabel = "Arch source"
      Write-Host "PROGRESS: Preparing Arch source for Xrefs transfer: $name"
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

    if ($stagedInXrefs) {
      $workingPath = New-IncomingWorkingCopy -SourcePath $sourcePath -TargetDir $targetDir -SourceLabel $stageLabel
      if ($archiveSelectedSource) {
        [void](Move-FileToArchive -FilePath $sourcePath -ArchiveRoot $targetDir -Message "Archived selected Xrefs source to")
      }
      Write-Host "PROGRESS: Final target in Xrefs: $([IO.Path]::GetFileName($targetPath))"
    }

    $existing = Get-ChildItem -LiteralPath $targetDir -File |
      Where-Object { $_.BaseName -ieq $targetName }
    Write-Host "PROGRESS: Checking exact target-name collisions for '$targetName' in $targetDir"
    foreach ($item in $existing) {
      if ($item.FullName -ne $workingPath) {
        [void](Move-FileToArchive -FilePath $item.FullName -ArchiveRoot $targetDir -Message "Archived existing file to")
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

if ($files.Count -gt 0) {
  $outputFolder = Split-Path -Parent ([IO.Path]::GetFullPath($files[0]))
  if (-not [string]::IsNullOrWhiteSpace($outputFolder)) {
    Write-Host "PROGRESS: OUTPUT_FOLDER: $outputFolder"
  }
}
