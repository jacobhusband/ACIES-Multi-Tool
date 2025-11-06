# detachSelectedUnderlays.ps1
# This script now correctly handles PDF, DGN, DWF, Raster Image, and DWG XREFs.
# 1. User selects DWG files.
# 2. accoreconsole runs a LISP to query the drawing database for all reference types.
# 3. PowerShell parses the file to get the reference names.
# 4. A GUI appears, letting the user select which references to detach.
# 5. The script generates the correct detachment commands for each reference type.
# 6. accoreconsole runs the new script, saves, and quits.
# 7. The script will pause at the end until the user presses Enter.

Write-Host "PROGRESS: Initializing script..."
Write-Host "=====================================================================" -ForegroundColor Yellow
Write-Host "IMPORTANT: Please close all AutoCAD and DWG files before proceeding!" -ForegroundColor Yellow
Write-Host "The script needs to write to these files and will fail if they are open." -ForegroundColor Yellow
Write-Host "====================================================================="

# --- CONFIGURATION ---
$acadCore = "C:\Program Files\Autodesk\AutoCAD 2025\accoreconsole.exe"
# --- END CONFIGURATION ---

# Relaunch in Single-Threaded Apartment (STA) mode for the GUI components
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $ps = (Get-Process -Id $PID).Path
    Start-Process -FilePath $ps -ArgumentList @("-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-File", "`"$PSCommandPath`"") -Wait
    exit
}

# --- VALIDATION ---
if (-not (Test-Path $acadCore)) { Write-Host "PROGRESS: ERROR: AutoCAD Core Console not found at $acadCore"; exit 1 }

# --- STAGE 1: Select DWG Files ---
Write-Host "PROGRESS: Waiting for DWG file selection..."
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
$dwgDialog = New-Object System.Windows.Forms.OpenFileDialog
$dwgDialog.Title = "Select DWG file(s) to process"
$dwgDialog.Filter = "DWG files (*.dwg)|*.dwg|All files (*.*)|*.*"
$dwgDialog.Multiselect = $true
$dwgDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
if ($dwgDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK -or -not $dwgDialog.FileNames) {
    Write-Host "PROGRESS: ERROR: No files selected."
    Write-Host "Press Enter to exit."; Read-Host | Out-Null; exit 1
}
$files = $dwgDialog.FileNames
Write-Host "PROGRESS: $($files.Count) DWG file(s) selected."

# --- STAGE 2: Scan DWGs for all Reference Names ---
Write-Host "PROGRESS: Scanning files to find all references. This may take a moment..."

# Define paths for our temporary files
$refListFile = Join-Path $env:TEMP "ref_list_temp.txt"
$lispScanFile = Join-Path $env:TEMP "scan_refs_temp.lsp"
$scriptScanFile = Join-Path $env:TEMP "run_scan_refs.scr"

# Clear any old results from previous runs
if (Test-Path $refListFile) { Remove-Item $refListFile }

# This LISP routine now scans for all reference types.
$lispScanContent = @"
(defun c:SCANREFS ( / f dictlist dict_info dictname prefix main_dict item block_data is_xref)
  (vl-load-com)
  (setq f (open "$($refListFile.replace('\', '/'))" "a"))

  ; --- Underlays and Images (Dictionary Method) ---
  (setq dictlist '(
    "ACAD_PDFDEFINITIONS:PDF"
    "ACAD_DGNDEFINITIONS:DGN"
    "ACAD_DWFDEFINITIONS:DWF"
    "ACAD_IMAGEDEF:IMAGE"
  ))
  (foreach dict_info dictlist
    (setq dictname (car (vl-string-split dict_info ":")))
    (setq prefix (cadr (vl-string-split dict_info ":")))
    (if (setq main_dict (dictsearch (namedobjdict) dictname))
      (foreach item main_dict
        (if (= (car item) 3) ; This is the definition name
          (write-line (strcat prefix ":" (cdr item)) f)
        )
      )
    )
  )

  ; --- DWG XREFs (Block Table Method) ---
  ; --- FIX --- This logic now correctly identifies XREFs and Overlays regardless of their "resolved" status.
  (while (setq block_data (tblnext "BLOCK" (not block_data)))
    (setq is_xref (cdr (assoc 70 block_data)))
    ; Check if it's an XREF (flag 4) or an XREF overlay (flag 8). This finds ALL XREFs.
    (if (or (= (logand is_xref 4) 4) (= (logand is_xref 8) 8))
      ; Exclude anonymous blocks (flag 1)
      (if (= (logand is_xref 1) 0)
        (write-line (strcat "DWG:" (cdr (assoc 2 block_data))) f)
      )
    )
  )
  (close f)
  (princ)
)
"@
Set-Content -Path $lispScanFile -Value $lispScanContent -Encoding ASCII

# This script loads the LISP file and calls the SCANREFS function.
$scriptScanContent = @"
(load "$($lispScanFile -replace '\\', '/')")
SCANREFS
QUIT
"@
Set-Content -Path $scriptScanFile -Value $scriptScanContent -Encoding ASCII

# Execute the scan on each DWG
Write-Host "--- AutoCAD Core Console Output (Scan Phase) ---"
foreach ($dwg in $files) {
    $name = [IO.Path]::GetFileName($dwg)
    Write-Host "PROGRESS: Scanning $name..."
    & $acadCore /i "$dwg" /s "$scriptScanFile" /l en-US 2>&1
}
Write-Host "--- End of AutoCAD Core Console Output ---"

# --- STAGE 3: User Selects References to Detach ---
if (-not (Test-Path $refListFile)) {
    Write-Host "PROGRESS: Scan completed, but no references were found. See console output above for details."
    Write-Host "Press Enter to exit."; Read-Host | Out-Null; exit 0
}

$uniqueRefs = Get-Content $refListFile | Sort-Object -Unique
if (-not $uniqueRefs) {
    Write-Host "PROGRESS: No references were found in the selected files."
    Write-Host "Press Enter to exit."; Read-Host | Out-Null; exit 0
}

# Create a more user-friendly GUI for selection
$selectedRefs = @()
$form = New-Object System.Windows.Forms.Form; $form.Text = "Select References to Detach"; $form.Size = New-Object System.Drawing.Size(400, 500); $form.StartPosition = "CenterScreen"
$label = New-Object System.Windows.Forms.Label; $label.Text = "Select references to detach (use Ctrl+Click or Shift+Click):"; $label.Location = New-Object System.Drawing.Point(10, 10); $label.AutoSize = $true; $form.Controls.Add($label)
$listBox = New-Object System.Windows.Forms.ListBox; $listBox.Location = New-Object System.Drawing.Point(10, 40); $listBox.Size = New-Object System.Drawing.Size(360, 350); $listBox.Items.AddRange($uniqueRefs)
$listBox.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
$form.Controls.Add($listBox)
$selectAllButton = New-Object System.Windows.Forms.Button; $selectAllButton.Text = "Select All"; $selectAllButton.Location = New-Object System.Drawing.Point(10, 400)
$selectAllButton.add_Click({ for ($i = 0; $i -lt $listBox.Items.Count; $i++) { $listBox.SetSelected($i, $true) } })
$form.Controls.Add($selectAllButton)
$deselectAllButton = New-Object System.Windows.Forms.Button; $deselectAllButton.Text = "Deselect All"; $deselectAllButton.Location = New-Object System.Drawing.Point(95, 400)
$deselectAllButton.add_Click({ $listBox.ClearSelected() })
$form.Controls.Add($deselectAllButton)
$okButton = New-Object System.Windows.Forms.Button; $okButton.Text = "OK"; $okButton.Location = New-Object System.Drawing.Point(285, 400); $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.AcceptButton = $okButton; $form.Controls.Add($okButton)
$form.Topmost = $true
$result = $form.ShowDialog()

if ($result -ne [System.Windows.Forms.DialogResult]::OK -or $listBox.SelectedItems.Count -eq 0) {
    Write-Host "PROGRESS: ERROR: No references selected for detachment. Operation cancelled."
    Write-Host "Press Enter to exit."; Read-Host | Out-Null; exit 1
}
$selectedRefs = $listBox.SelectedItems
Write-Host "PROGRESS: $($selectedRefs.Count) reference(s) selected for detachment."


# --- STAGE 4: Build Detach Script and Process DWGs ---

# Sort selected items into different lists based on their type.
$underlaysAndImages = @()
$dwgsToDetach = @()

foreach ($item in $selectedRefs) {
    $parts = $item.Split(':', 2)
    $type = $parts[0]
    $name = $parts[1]

    if ($type -eq 'DWG') {
        $dwgsToDetach += $name
    } else {
        # This list will contain PDFs, DGNs, DWFs, Images
        $underlaysAndImages += $name
    }
}

# --- Build DWG Detach Commands ---
$dwgDetachCommands = ""
if ($dwgsToDetach.Count -gt 0) {
    foreach ($dwgName in $dwgsToDetach) {
        $dwgDetachCommands += "_.-XREF`nD`n`"$dwgName`"`n"
    }
}

# --- Build Underlay/Image Detach LISP ---
$lispDetachFile = Join-Path $env:TEMP "detach_others_temp.lsp"
$lispDetachContent = ""
if ($underlaysAndImages.Count -gt 0) {
    $lispDetachTemplate = @"
(defun c:DETACHOTHERS ( / namelist dictlist dictname dictobj refname ss)
  (setq namelist namelist_placeholder)
  (foreach refname namelist
    (setq ss (ssget "_X" (list (cons 2 refname))))
    (if ss (command "_.ERASE" ss ""))
  )
  (setq dictlist '("ACAD_PDFDEFINITIONS" "ACAD_DGNDEFINITIONS" "ACAD_DWFDEFINITIONS" "ACAD_IMAGEDEF"))
  (foreach dictname dictlist
    (if (setq dictobj (dictsearch (namedobjdict) dictname))
      (foreach refname namelist
        (if (dictsearch (cdr (assoc -1 dictobj)) refname)
          (dictremove (cdr (assoc -1 dictobj)) refname)
        )
      )
    )
  )
  (princ)
)
"@
    $lispListString = "'(" + '"' + ($underlaysAndImages -join '" "') + '"' + ")"
    $lispDetachContent = $lispDetachTemplate.Replace("namelist_placeholder", $lispListString)
    Set-Content -Path $lispDetachFile -Value $lispDetachContent -Encoding ASCII
}

# --- Combine into Final Script ---
$scriptDetachFile = Join-Path $env:TEMP "run_DETACH.scr"
$lispDetachPathForScript = ($lispDetachFile -replace '\\', '/')

# Start with the DWG detach commands
$scriptDetachContent = $dwgDetachCommands

# If there are other types, add the LISP commands
if ($underlaysAndImages.Count -gt 0) {
    $scriptDetachContent += @"
(load "$lispDetachPathForScript")
DETACHOTHERS
"@
}

# Add the final save and quit commands
$scriptDetachContent += @"
QSAVE
QUIT
"@

Set-Content -Path $scriptDetachFile -Value $scriptDetachContent -Encoding ASCII

# Process the files
$failed = @()
$i = 0
Write-Host "--- AutoCAD Core Console Output (Detach Phase) ---"
foreach ($dwg in $files) {
    $i++; $name = [IO.Path]::GetFileName($dwg)
    Write-Host "PROGRESS: Processing $i of $($files.Count): $name"
    & $acadCore /i "$dwg" /s "$scriptDetachFile" 2>&1
    if ($LASTEXITCODE -ne 0) { $failed += $dwg }
}
Write-Host "--- End of AutoCAD Core Console Output ---"


# --- STAGE 5: Final Report and Cleanup ---
if ($failed.Count) { Write-Host "PROGRESS: ERROR: $($failed.Count) of $($files.Count) file(s) failed to process." }
else { Write-Host "PROGRESS: Successfully processed $($files.Count) drawing(s)." }

# Clean up all temporary files
Remove-Item $refListFile, $lispScanFile, $scriptScanFile, $scriptDetachFile, $lispDetachFile -ErrorAction SilentlyContinue

Write-Host "PROGRESS: Done."
Write-Host ""; Write-Host "Script has finished. Press Enter to close this window."; Read-Host | Out-Null