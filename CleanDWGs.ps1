param(
    [string]$TitleblockPath,
    [string]$Disciplines,
    [string]$OutputFolder,
    [string]$ProjectRoot,
    [string]$AcadPath
)

#--- WINDOW MANAGEMENT ---
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")]
    public static extern bool FlashWindow(IntPtr hWnd, bool bInvert);
    public const int SW_RESTORE = 9;
}
"@

function Bring-Acad-To-Front {
    param($AcadApp)
    try {
        $hwndWait = 0; $maxHwndWait = 10
        while ($null -eq $AcadApp.HWND -and $hwndWait -lt $maxHwndWait) {
            Start-Sleep -Seconds 1; $hwndWait++
        }
        
        $hWnd = $AcadApp.HWND
        if ($hWnd -ne $null) {
            [Win32]::ShowWindow($hWnd, [Win32]::SW_RESTORE)
            Start-Sleep -Milliseconds 200
            [Win32]::SetForegroundWindow($hWnd)
            Start-Sleep -Milliseconds 100
            [Win32]::FlashWindow($hWnd, $true)
        }
    }
    catch {
        Write-Log "Could not bring AutoCAD to front: $($_.Exception.Message)" "WARNING"
    }
}

#--- LOGGING & USER PROMPTS ---
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "HH:mm:ss"
    $prefix = if ($Level -eq "ERROR") { "ERROR" } elseif ($Level -eq "WARNING") { "WARN" } else { "INFO" }
    Write-Host "PROGRESS: [$timestamp] $prefix - $Message"
}

function Show-UserPrompt {
    param([string]$Message, [string]$Title = "Review Required", [System.Windows.Forms.MessageBoxButtons]$Buttons)
    Add-Type -AssemblyName System.Windows.Forms
    $result = [System.Windows.Forms.MessageBox]::Show(
        $Message, $Title, $Buttons,
        [System.Windows.Forms.MessageBoxIcon]::Question,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button1,
        [System.Windows.Forms.MessageBoxOptions]::DefaultDesktopOnly
    )
    return $result
}

#--- ABORT SIGNAL CHECK ---
function Test-AbortSignal {
    $abortFile = Join-Path $env:TEMP "abort_cleandwgs.flag"
    if (Test-Path $abortFile) {
        Write-Log "Abort signal detected. Stopping..." "WARNING"
        Remove-Item $abortFile -Force -ErrorAction SilentlyContinue
        return $true
    }
    return $false
}

#--- OPTIMIZED WAIT FUNCTIONS ---
function Wait-For-AutoCAD-Ready {
    param($AcadApp, [int]$TimeoutSeconds = 45, [switch]$ExtraWait)
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    while ($stopwatch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        try {
            if ($AcadApp.GetAcadState().IsQuiescent -and ($AcadApp.Documents.Count -eq 0 -or $AcadApp.ActiveDocument.GetVariable("CMDACTIVE") -eq 0)) {
                # Base wait to ensure stability
                Start-Sleep -Milliseconds 500
                
                # Extra wait if document just opened to ensure full load
                if ($ExtraWait) {
                    Start-Sleep -Milliseconds 1500
                }
                return $true
            }
        }
        catch { }
        Start-Sleep -Milliseconds 250
    }
    throw "Timeout: AutoCAD did not become ready within $TimeoutSeconds seconds."
}

function Wait-For-Command-Complete {
    param($AcadApp, [int]$TimeoutSeconds = 300)
    Write-Log "Waiting for current command to complete..."
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $lastMessageTime = $stopwatch.Elapsed.TotalSeconds

    while ($stopwatch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        if (Test-AbortSignal) { throw "Operation aborted by user" }
        
        try {
            if ($AcadApp.GetAcadState().IsQuiescent -and $AcadApp.ActiveDocument.GetVariable("CMDACTIVE") -eq 0) {
                Start-Sleep -Milliseconds 250
                Write-Log "Command completed."
                return $true
            }
        }
        catch { }
        
        if (($stopwatch.Elapsed.TotalSeconds - $lastMessageTime) -ge 10) {
            Write-Log "Still waiting for command to complete..."
            $lastMessageTime = $stopwatch.Elapsed.TotalSeconds
        }
        
        Start-Sleep -Milliseconds 250
    }
    throw "Command timeout after $TimeoutSeconds seconds. AutoCAD may be waiting for user input."
}

#--- ROBUST AUTOCAD INITIALIZATION ---
function Initialize-AutoCAD {
    # Determine preferred progId from AcadPath
    $preferredProgId = $null
    if ($AcadPath) {
        if ($AcadPath -match "AutoCAD (\d{4})") {
            $year = $matches[1]
            $suffix = ($year - 2000).ToString()
            $preferredProgId = "AutoCAD.Application.$suffix"
        }
    }
    $progIds = @()
    if ($preferredProgId) { $progIds += $preferredProgId }
    $progIds += @("AutoCAD.Application.25", "AutoCAD.Application.24", "AutoCAD.Application.23", "AutoCAD.Application.22", "AutoCAD.Application")
    $progIds = $progIds | Select-Object -Unique
    $script:AcadWasStarted = $false
    
    Write-Log "Searching for a running instance of AutoCAD (2025 preferred)..."
    foreach ($progId in $progIds) {
        try {
            $Acad = [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId)
            if ($Acad) {
                Write-Log "Successfully connected to running instance of $progId."
                $script:AcadWasStarted = $false
                $Acad.Visible = $true
                break
            }
        }
        catch { }
    }
    
    if ($null -eq $Acad) {
        Write-Log "No running instance found. Starting a new AutoCAD process (2025 preferred)..."
        foreach ($progId in $progIds) {
            try {
                $Acad = New-Object -ComObject $progId
                if ($Acad) {
                    Write-Log "AutoCAD process started ($progId). Waiting for it to initialize..."
                    $script:AcadWasStarted = $true
                    $Acad.Visible = $true
                    
                    $initStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                    while ($initStopwatch.Elapsed.TotalSeconds -lt 60) {
                        try {
                            $test = $Acad.Name
                            Write-Log "$progId has initialized successfully."
                            break # Exit inner loop
                        }
                        catch { Start-Sleep -Milliseconds 250 }
                    }
                    if ($Acad) { break } # Exit outer loop
                }
            }
            catch { }
        }
    }
    
    if ($null -eq $Acad) {
        throw "Failed to connect to or start AutoCAD. Please ensure AutoCAD is installed correctly."
    }

    Write-Log "Setting application-level silent mode to prevent pop-ups..."
    try {
        $Acad.SetVariable("FILEDIA", 0)
        $Acad.SetVariable("CMDDIA", 0)
        $Acad.SetVariable("PROXYNOTICE", 0)
        $Acad.SetVariable("XREFNOTIFY", 0)
        Write-Log "Application prepared for silent operation."
    }
    catch {
        Write-Log "Could not set initial silent mode. Pop-ups may occur. Error: $($_.Exception.Message)" "WARNING"
    }

    if ($script:AcadWasStarted) {
        for ($i = $Acad.Documents.Count - 1; $i -ge 0; $i--) {
            $doc = $Acad.Documents.Item($i)
            if ($doc.Name -like "Drawing*.dwg") {
                try {
                    Write-Log "Closing default drawing: $($doc.Name)"
                    $doc.Close($false)
                }
                catch {}
            }
        }
    }
    
    return $Acad
}

function Show-ManualClosePrompt {
    param($AcadApp, $DocFullName)
    $message = "Please close '$(Split-Path $DocFullName -Leaf)' manually in AutoCAD when you are finished, then click OK here to continue the process."
    Show-UserPrompt -Message $message -Title "Manual Action Required" -Buttons ([System.Windows.Forms.MessageBoxButtons]::OK)
    
    Write-Log "Waiting for user to manually close '$((Split-Path $DocFullName -Leaf))'..."
    while ($AcadApp.Documents | Where-Object { $_.FullName -eq $DocFullName }) {
        if (Test-AbortSignal) { throw "Operation aborted during manual close wait" }
        Start-Sleep -Seconds 2
    }
    Write-Log "Document closed by user. Proceeding..."
}

#--- PROCESS TITLEBLOCK (REVISED WORKFLOW) ---
function Process-Titleblock {
    param($AcadApp, $TbPath)
    
    Write-Log "=== 1. Processing Titleblock ==="
    $tbName = [System.IO.Path]::GetFileName($TbPath)
    
    if (Test-AbortSignal) { throw "Operation aborted by user" }
    
    $tbDoc = $null
    try {
        $tbDoc = $AcadApp.Documents.Open($TbPath)
    }
    catch {
        $errorMsg = "Failed to open the title block file: '$TbPath'.`n`n"
        $errorMsg += "Possible causes: File is corrupt, not a valid DWG, or a permissions issue exists.`n`n"
        $errorMsg += "Original Error: $($_.Exception.Message)"
        throw $errorMsg
    }
    
    Wait-For-AutoCAD-Ready -AcadApp $AcadApp -ExtraWait
    Bring-Acad-To-Front -AcadApp $AcadApp
    
    Write-Log "Titleblock '$tbName' loaded and ready for cleaning."
    
    # Show instruction prompt - use TopMost to ensure visibility
    Add-Type -AssemblyName System.Windows.Forms
    $msgBox = [System.Windows.Forms.MessageBox]::Show(
        "The title block '$tbName' is now open in AutoCAD.`n`n" +
        "Please select the two corner points of the title block when prompted.`n`n" +
        "Click OK to start the cleaning process.",
        "Action Required",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
    
    # Check if user clicked OK or closed the dialog
    if ($msgBox -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Log "User cancelled the operation." "WARNING"
        try { $tbDoc.Close($false) } catch {}
        return $false
    }
    
    Write-Log "Sending CLEANTBLK command..."
    $tbDoc.SendCommand("._CLEANTBLK`n")
    
    Write-Log "Waiting for CLEANTBLK command to complete (user interaction required)..."
    Wait-For-Command-Complete -AcadApp $AcadApp -TimeoutSeconds 300
    
    Write-Log "CLEANTBLK command completed. Awaiting review..."
    
    # Review prompt
    $reviewMessage = @"
Please review the cleaned title block.

What would you like to do?

- YES: Cleaning looks good. Save the changes and continue.
- NO: Undo the cleaning. I will handle it manually.
- CANCEL: Leave the cleaning as is, but let me fix it manually.
"@

    $choice = [System.Windows.Forms.MessageBox]::Show(
        $reviewMessage,
        "Titleblock Review",
        [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
        [System.Windows.Forms.MessageBoxIcon]::Question,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button1
    )
    
    switch ($choice) {
        "Yes" {
            Write-Log "User choice: 'Cleaning looks good, please continue'."
            $tbDoc.Save()
            $tbDoc.Close()
            return $true
        }
        "No" {
            Write-Log "User choice: 'Undo the cleaning, I will handle it'."
            $tbDoc.SendCommand("._U`n")
            Wait-For-Command-Complete -AcadApp $AcadApp -TimeoutSeconds 30
            Show-ManualClosePrompt -AcadApp $AcadApp -DocFullName $tbDoc.FullName
            return $true
        }
        "Cancel" {
            Write-Log "User choice: 'Leave the cleaning, but let me fix it'."
            Show-ManualClosePrompt -AcadApp $AcadApp -DocFullName $tbDoc.FullName
            return $true
        }
        Default {
            Write-Log "User cancelled the operation during titleblock review. Halting process." "WARNING"
            try { $tbDoc.Close($false) } catch {}
            return $false
        }
    }
}

#--- PROCESS CAD FILE ---
function Process-CADFile {
    param($AcadApp, $DwgPath, $OutputRoot)
    
    $dwgName = [System.IO.Path]::GetFileName($DwgPath)
    Write-Log "Processing: $dwgName"
    
    if (Test-AbortSignal) { throw "Operation aborted by user" }
    if (-not (Test-Path $DwgPath)) { throw "File not found: $DwgPath" }
    
    $doc = $null
    try {
        $doc = $AcadApp.Documents.Open($DwgPath)
        
        # Wait for document to fully load with extra time
        Wait-For-AutoCAD-Ready -AcadApp $AcadApp -ExtraWait
        
        # Activate the document and wait again
        $doc.Activate()
        Start-Sleep -Milliseconds 500
        
        Write-Log "Executing CLEANCAD command..."
        $doc.SendCommand("._CLEANCAD`n")
        Wait-For-Command-Complete -AcadApp $AcadApp -TimeoutSeconds 180
        
        # Additional wait after command completes to ensure all operations finish
        Start-Sleep -Milliseconds 1000
        
        $layoutNames = @()
        foreach ($layout in $doc.Layouts) {
            if ($layout.Name.ToUpper() -ne "MODEL") { $layoutNames += $layout.Name }
        }
        
        if ($layoutNames.Count -gt 0) {
            # Save to project root with layout names
            $newName = ($layoutNames -join " ") + ".dwg"
            $newPath = Join-Path $OutputRoot $newName
            Write-Log "Saving cleaned file to project root: '$newName'..."
            $doc.SaveAs($newPath)
        }
        else {
            # No layouts found, save with original name to project root
            $originalName = [System.IO.Path]::GetFileName($DwgPath)
            $newPath = Join-Path $OutputRoot $originalName
            Write-Log "No layouts found. Saving to project root as '$originalName'..."
            $doc.SaveAs($newPath)
        }
        
        # Wait after save to ensure it completes
        Start-Sleep -Milliseconds 500
        
        $doc.Close()
        
        # Wait after close to ensure clean shutdown
        Start-Sleep -Milliseconds 500
        
        return $true
    }
    catch {
        Write-Log "Error processing '$dwgName`: $($_.Exception.Message)" "ERROR"
        if ($doc) { try { $doc.Close($false) } catch { } }
        throw
    }
}

#--- MAIN EXECUTION ---
try {
    Write-Log "=== CLEAN DWGS SCRIPT STARTED ==="
    Write-Log "Output Folder: $OutputFolder"
    
    if (-not $TitleblockPath -or -not (Test-Path $TitleblockPath)) {
        throw "Titleblock DWG not found at path: $TitleblockPath"
    }
    
    $DisciplineList = @()
    if (-not [string]::IsNullOrWhiteSpace($Disciplines)) {
        $DisciplineList = $Disciplines -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
    }
    
    $Acad = Initialize-AutoCAD
    
    $allDwgFiles = @()
    Write-Log "Scanning for CAD drawings to process..."
    foreach ($discipline in $DisciplineList) {
        $disciplinePath = Join-Path $OutputFolder $discipline
        if (Test-Path $disciplinePath) {
            Write-Log "Scanning path: '$disciplinePath'"
            $dwgFilesInDiscipline = Get-ChildItem -Path $disciplinePath -Filter "*.dwg" -File -Recurse | Where-Object { $_.FullName -ne $TitleblockPath }
            $count = ($dwgFilesInDiscipline | Measure-Object).Count
            Write-Log "Found $count DWG(s) in '$discipline' to process."
            if ($count -gt 0) {
                $allDwgFiles += $dwgFilesInDiscipline.FullName
            }
        }
        else {
            Write-Log "WARNING: Discipline folder not found, skipping scan: $disciplinePath" "WARNING"
        }
    }
    
    $allDwgFiles = $allDwgFiles | Select-Object -Unique
    $totalFiles = $allDwgFiles.Count
    Write-Log "Found $totalFiles total unique CAD files to process (excluding the titleblock)."
    
    $continueProcessing = Process-Titleblock -AcadApp $Acad -TbPath $TitleblockPath
    
    if (-not $continueProcessing) {
        throw "Operation halted by user after titleblock processing."
    }
    
    Write-Log "=== 2. Processing CAD Drawings ==="
    if ($totalFiles -eq 0) {
        Write-Log "No additional CAD files were found in the selected discipline folders to process."
    }
    else {
        $processedCount = 0; $failedCount = 0
        foreach ($dwgPath in $allDwgFiles) {
            $processedCount++
            try {
                Write-Log "--- File $processedCount of $totalFiles ---"
                Process-CADFile -AcadApp $Acad -DwgPath $dwgPath -OutputRoot $OutputFolder
            }
            catch {
                $failedCount++
                $errorMsg = $_.Exception.Message
                Write-Log "ERROR: Failed to process $(Split-Path $dwgPath -Leaf): $errorMsg" "ERROR"
                if ($errorMsg -like "*aborted by user*") { throw $_ }
                continue
            }
        }
        
        Write-Log "=== PROCESSING COMPLETE ==="
        Write-Log "Successfully processed: $($processedCount - $failedCount) files."
        if ($failedCount -gt 0) {
            Write-Log "Failed to process: $failedCount files." "WARNING"
        }
    }
    
    Write-Log "Opening output folder for review..."
    if (Test-Path $OutputFolder) {
        Start-Process explorer.exe -ArgumentList "`"$OutputFolder`""
    }
    
    Write-Log "=== SCRIPT FINISHED SUCCESSFULLY ==="
}
catch {
    $errorMsg = $_.Exception.Message
    $errorLevel = if ($errorMsg -like "*aborted by user*" -or $errorMsg -like "*halted by user*") { "WARNING" } else { "ERROR" }
    Write-Log "Script terminated: $errorMsg" $errorLevel
    
    if ($errorLevel -eq "ERROR") {
        Show-UserPrompt -Message "An unexpected error occurred:`n`n$errorMsg`n`nPlease check the application log for details." -Title "Clean DWGs Tool Error" -Buttons ([System.Windows.Forms.MessageBoxButtons]::OK)
    }
}
finally {
    if ($Acad -and $script:AcadWasStarted) {
        try {
            Write-Log "Closing AutoCAD instance started by the script..."
            $Acad.Quit()
        }
        catch { }
    }
    
    $abortFile = Join-Path $env:TEMP "abort_cleandwgs.flag"
    if (Test-Path $abortFile) {
        Remove-Item $abortFile -Force -ErrorAction SilentlyContinue
    }
}