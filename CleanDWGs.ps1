param(
    [string]$TitleblockPath,
    [string]$Disciplines,
    [string]$OutputFolder,
    [string]$ProjectRoot
)

#--- WINDOW MANAGEMENT (from old version) ---
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
        # Wait for HWND to be available
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
    param([string]$Message, [string]$Title = "Review Required")
    Add-Type -AssemblyName System.Windows.Forms
    $result = [System.Windows.Forms.MessageBox]::Show(
        $Message, $Title,
        [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
        [System.Windows.Forms.MessageBoxIcon]::Question,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button1,
        [System.Windows.Forms.MessageBoxOptions]::DefaultDesktopOnly
    )
    return $result  # Returns: Yes, No, or Cancel
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

#--- ENHANCED WAIT FUNCTIONS ---
function Wait-For-AutoCAD-Ready {
    param($AcadApp, [int]$TimeoutSeconds = 30)
    $elapsed = 0
    while ($elapsed -lt $TimeoutSeconds) {
        try {
            $state = $AcadApp.GetAcadState()
            $cmdActive = $AcadApp.ActiveDocument.GetVariable("CMDACTIVE")
            if ($state.IsQuiescent -and $cmdActive -eq 0) {
                Start-Sleep -Milliseconds 500  # Stabilization delay
                return $true
            }
        }
        catch { }
        Start-Sleep -Milliseconds 500
        $elapsed += 0.5
    }
    Write-Log "WARNING: AutoCAD readiness timeout after $TimeoutSeconds seconds." "WARNING"
}

function Wait-For-Command-Complete {
    param($AcadApp, [int]$TimeoutSeconds = 120)
    $elapsed = 0
    Write-Log "Waiting for command to complete..."
    while ($elapsed -lt $TimeoutSeconds) {
        if (Test-AbortSignal) { throw "Operation aborted by user" }
        
        try {
            $state = $AcadApp.GetAcadState()
            $cmdActive = $AcadApp.ActiveDocument.GetVariable("CMDACTIVE")
            if ($state.IsQuiescent -and $cmdActive -eq 0) {
                Start-Sleep -Milliseconds 500  # Ensure stability
                Write-Log "Command completed successfully."
                return $true
            }
        }
        catch { }
        
        Start-Sleep -Milliseconds 500
        $elapsed += 0.5
    }
    throw "Command timeout after $TimeoutSeconds seconds. AutoCAD may be waiting for input."
}

#--- ROBUST AUTOCAD INITIALIZATION (from old version) ---
function Initialize-AutoCAD {
    $progIds = @("AutoCAD.Application", "AutoCAD.Application.24", "AutoCAD.Application.23", "AutoCAD.Application.22")
    $maxRetries = 3; $retryCount = 0
    $script:AcadWasStarted = $false
    
    while ($null -eq $Acad -and $retryCount -lt $maxRetries) {
        $retryCount++
        
        # Try connecting to running instance
        foreach ($progId in $progIds) {
            try {
                $Acad = [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId)
                if ($Acad) {
                    Write-Log "Connected to running AutoCAD ($progId)"
                    $script:AcadWasStarted = $false
                    break
                }
            }
            catch { }
        }
        
        # If no running instance, start new one
        if ($null -eq $Acad) {
            Write-Log "No running AutoCAD found. Starting new instance..."
            foreach ($progId in $progIds) {
                try {
                    $Acad = New-Object -ComObject $progId
                    $script:AcadWasStarted = $true
                    $Acad.Visible = $true
                    
                    # Wait for initialization
                    $initWait = 0; $maxInitWait = 30
                    while ($initWait -lt $maxInitWait) {
                        try { $test = $Acad.Application; break }
                        catch { Start-Sleep -Seconds 1; $initWait++ }
                    }
                    Write-Log "AutoCAD started and initialized ($progId)"
                    break
                }
                catch { }
            }
        }
        
        if ($null -eq $Acad -and $retryCount -lt $maxRetries) {
            Write-Log "Retrying in 5 seconds..."
            Start-Sleep -Seconds 5
        }
    }
    
    if ($null -eq $Acad) {
        throw "Failed to connect to or start AutoCAD after $maxRetries attempts."
    }
    
    # Clean up default drawings if we started AutoCAD
    if ($script:AcadWasStarted) {
        $defaultDocNames = @("Drawing1.dwg", "Drawing2.dwg")
        $docsToClose = $Acad.Documents | Where-Object { $defaultDocNames -contains $_.Name }
        foreach ($doc in $docsToClose) {
            try { $doc.Close($false); Wait-For-AutoCAD-Ready -AcadApp $Acad }
            catch { }
        }
    }
    
    return $Acad
}

#--- PROCESS TITLEBLOCK ---
function Process-Titleblock {
    param($AcadApp, $TbPath)
    
    Write-Log "=== Processing Titleblock ==="
    $tbName = [System.IO.Path]::GetFileName($TbPath)
    Write-Log "Opening: $tbName"
    
    if (Test-AbortSignal) { throw "Operation aborted by user" }
    
    $tbDoc = $AcadApp.Documents.Open($TbPath)
    Set-SilentMode -Doc $tbDoc
    Wait-For-AutoCAD-Ready -AcadApp $AcadApp
    Bring-Acad-To-Front -AcadApp $AcadApp
    
    Write-Log "Executing CLEANTBLK command..."
    $tbDoc.SendCommand("._CLEANTBLK`n")
    
    # CRITICAL: Wait for command to truly complete
    Write-Log "Please complete the CLEANTBLK command by selecting the titleblock corners in AutoCAD."
    Wait-For-Command-Complete -AcadApp $AcadApp -TimeoutSeconds 300
    
    # Now prompt for user action
    $message = @"
Cleaning of titleblock '$tbName' is complete.

Please review the drawing in AutoCAD.

- Click YES to SAVE the changes and continue.
- Click NO to UNDO the changes (then close manually).
- Click CANCEL to LEAVE the changes and close manually.
"@
    
    $choice = Show-UserPrompt -Message $message -Title "Titleblock Review"
    
    switch ($choice) {
        "Yes" {
            Write-Log "User chose: Save and Continue"
            $tbDoc.Save()
            $tbDoc.Close()
            return $true
        }
        "No" {
            Write-Log "User chose: Undo and Handle Manually"
            $tbDoc.SendCommand("._UNDO`n1`n")
            Show-ManualClosePrompt -AcadApp $AcadApp -FileName $tbName
            return $true
        }
        "Cancel" {
            Write-Log "User chose: Leave and Handle Manually"
            Show-ManualClosePrompt -AcadApp $AcadApp -FileName $tbName
            return $true
        }
    }
    return $false
}

function Show-ManualClosePrompt {
    param($AcadApp, $FileName)
    $message = "Please close '$FileName' manually in AutoCAD when ready, then click OK to continue."
    [System.Windows.Forms.MessageBox]::Show($message, "Manual Close Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    
    # Wait for file to be closed
    while ($AcadApp.Documents | Where-Object { $_.Name -eq $FileName }) {
        if (Test-AbortSignal) { throw "Operation aborted during manual close wait" }
        Start-Sleep -Seconds 2
    }
}

#--- SET SILENT MODE ---
function Set-SilentMode {
    param($Doc)
    try {
        $variables = @("FILEDIA=0", "EXPERT=5", "XREFNOTIFY=0", "OLELINKSDIALOG=0", "XLOADCTL=0", "PROXYNOTICE=0", "ACADLSPASDOC=0", "CMDDIA=0", "ATTDIA=0")
        foreach ($var in $variables) {
            try { $Doc.SendCommand("._SETVAR`n$var`n") } catch { }
        }
        try { $Doc.SendCommand("._EXCELWARNINGS 0`n") } catch { }
        Write-Log "Silent mode configured"
    }
    catch {
        Write-Log "Warning setting silent mode: $($_.Exception.Message)" "WARNING"
    }
}

#--- PROCESS CAD FILE ---
function Process-CADFile {
    param($AcadApp, $DwgPath)
    
    $dwgName = [System.IO.Path]::GetFileName($DwgPath)
    Write-Log "=== Processing: $dwgName ==="
    
    if (Test-AbortSignal) { throw "Operation aborted by user" }
    if (-not (Test-Path $DwgPath)) { throw "File not found: $DwgPath" }
    
    $doc = $null
    try {
        $doc = $AcadApp.Documents.Open($DwgPath)
        Set-SilentMode -Doc $doc
        $doc.Activate()
        Bring-Acad-To-Front -AcadApp $AcadApp
        Wait-For-AutoCAD-Ready -AcadApp $AcadApp
        
        Write-Log "Executing CLEANCAD command..."
        $doc.SendCommand("._CLEANCAD`n")
        Wait-For-Command-Complete -AcadApp $AcadApp -TimeoutSeconds 180
        
        # Get layout names for renaming
        $layoutNames = @()
        foreach ($layout in $doc.Layouts) {
            if ($layout.Name -ne "Model") { $layoutNames += $layout.Name }
        }
        
        $message = @"
Processing of '$dwgName' is complete.

- Click YES to SAVE as '$($layoutNames -join " ").dwg' and continue.
- Click NO to UNDO changes (then close manually).
- Click CANCEL to LEAVE changes and close manually.
"@
        
        $choice = Show-UserPrompt -Message $message -Title "CAD File Review"
        
        switch ($choice) {
            "Yes" {
                if ($layoutNames.Count -gt 0) {
                    $newName = ($layoutNames -join " ") + ".dwg"
                    $newPath = Join-Path (Split-Path $DwgPath -Parent) $newName
                    $doc.SaveAs($newPath)
                    Write-Log "Saved as: $newName"
                }
                else {
                    $doc.Save()
                }
                $doc.Close()
            }
            "No" {
                $doc.SendCommand("._UNDO`n1`n")
                Show-ManualClosePrompt -AcadApp $AcadApp -FileName $dwgName
            }
            "Cancel" {
                Show-ManualClosePrompt -AcadApp $AcadApp -FileName $dwgName
            }
        }
        return $true
    }
    catch {
        Write-Log "Error processing $dwgName`: $($_.Exception.Message)" "ERROR"
        if ($doc) { try { $doc.Close($false) } catch { } }
        throw
    }
}

#--- MAIN EXECUTION ---
try {
    Write-Log "=== CLEAN DWGS SCRIPT STARTED ==="
    Write-Log "Parameters: TitleblockPath=$TitleblockPath, Disciplines=$Disciplines"
    Write-Log "OutputFolder=$OutputFolder, ProjectRoot=$ProjectRoot"
    
    if (-not $TitleblockPath -or -not (Test-Path $TitleblockPath)) {
        throw "Titleblock DWG not found at path: $TitleblockPath"
    }
    
    $DisciplineList = @()
    if (-not [string]::IsNullOrWhiteSpace($Disciplines)) {
        $DisciplineList = $Disciplines -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
    }
    
    Write-Log "Processing discipline folders: $($DisciplineList -join ', ')"
    
    $Acad = Initialize-AutoCAD
    Wait-For-AutoCAD-Ready -AcadApp $Acad
    
    # --- DISCOVER ALL CAD FILES ---
    $allDwgFiles = @()
    foreach ($discipline in $DisciplineList) {
        $disciplinePath = Join-Path $OutputFolder $discipline
        if (Test-Path $disciplinePath) {
            Write-Log "Scanning $discipline folder..."
            
            # Get all DWGs except titleblock and Xref folder contents (except titleblock itself)
            $dwgFiles = Get-ChildItem -Path $disciplinePath -Filter "*.dwg" -File -Recurse | 
            Where-Object { 
                $_.FullName -ne $TitleblockPath -and 
                -not ($_.DirectoryName -like "*\Xrefs\*" -and $_.FullName -ne $TitleblockPath)
            } | 
            ForEach-Object { $_.FullName }
            
            Write-Log "Found $($dwgFiles.Count) DWG files in $discipline"
            $allDwgFiles += $dwgFiles
        }
        else {
            Write-Log "WARNING: Discipline folder not found: $disciplinePath" "WARNING"
        }
    }
    
    $allDwgFiles = $allDwgFiles | Select-Object -Unique
    $totalFiles = $allDwgFiles.Count
    Write-Log "Total CAD files to process: $totalFiles"
    
    if ($totalFiles -eq 0) {
        Write-Log "WARNING: No CAD files found to process" "WARNING"
    }
    
    # --- PROCESS TITLEBLOCK ---
    $continueProcessing = Process-Titleblock -AcadApp $Acad -TbPath $TitleblockPath
    
    if (-not $continueProcessing) {
        Write-Log "Titleblock processing was cancelled. Skipping CAD files."
    }
    else {
        # --- PROCESS CAD FILES ---
        $processedCount = 0; $failedCount = 0
        
        foreach ($dwgPath in $allDwgFiles) {
            $processedCount++
            try {
                Write-Log "=== File $processedCount of $totalFiles ==="
                Process-CADFile -AcadApp $Acad -DwgPath $dwgPath
            }
            catch {
                $failedCount++
                $errorMsg = $_.Exception.Message
                Write-Log "ERROR: Failed to process $(Split-Path $dwgPath -Leaf): $errorMsg" "ERROR"
                
                if ($errorMsg -eq "Operation aborted by user") {
                    throw $_  # Re-throw to abort entire process
                }
                continue  # Continue to next file
            }
        }
        
        Write-Log "=== PROCESSING COMPLETE ==="
        Write-Log "Successfully processed: $($processedCount - $failedCount) files"
        if ($failedCount -gt 0) {
            Write-Log "Failed: $failedCount files" "WARNING"
        }
    }
    
    # --- OPEN OUTPUT FOLDER ---
    Write-Log "Opening output folder..."
    if (Test-Path $OutputFolder) {
        Start-Process explorer.exe -ArgumentList "`"$OutputFolder`"" -ErrorAction SilentlyContinue
        Write-Log "Output folder opened."
    }
    else {
        throw "Output folder not found at $OutputFolder"
    }
    
    Write-Log "=== SCRIPT FINISHED SUCCESSFULLY ==="
}
catch {
    $errorMsg = $_.Exception.Message
    $errorLevel = if ($errorMsg -eq "Operation aborted by user") { "WARNING" } else { "ERROR" }
    Write-Log "Script terminated: $errorMsg" $errorLevel
    
    if ($errorMsg -ne "Operation aborted by user") {
        [System.Windows.Forms.MessageBox]::Show(
            "An error occurred:`n`n$errorMsg`n`nPlease check the output for details.",
            "Clean DWGs Tool Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}
finally {
    if ($Acad) {
        try {
            foreach ($doc in $Acad.Documents) {
                try { $doc.Close($false) } catch { }
            }
            if ($script:AcadWasStarted) {
                $Acad.Quit()
            }
        }
        catch { }
    }
    
    $abortFile = Join-Path $env:TEMP "abort_cleandwgs.flag"
    if (Test-Path $abortFile) {
        Remove-Item $abortFile -Force -ErrorAction SilentlyContinue
    }
}