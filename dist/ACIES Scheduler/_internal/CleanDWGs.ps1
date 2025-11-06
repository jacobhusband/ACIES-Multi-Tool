param(
    [string]$TitleblockPath,
    [string[]]$CadDwgPaths,
    [string]$TitleblockSize,
    [string]$OutputFolder
)

# --- SCRIPT INITIALIZATION ---
function Write-Log {
    param ([string]$Message)
    $timestamp = Get-Date -Format "HH:mm:ss"
    $logMessage = "[$timestamp] $Message"
    Write-Host $logMessage
    Write-Host "PROGRESS: $logMessage"
}

function Show-UserPrompt {
    param(
        [string]$Message,
        [string]$Title = "User Action Required"
    )
    Add-Type -AssemblyName System.Windows.Forms
    $ButtonType = [System.Windows.Forms.MessageBoxButtons]::YesNoCancel
    $IconType = [System.Windows.Forms.MessageBoxIcon]::Question
    $Result = [System.Windows.Forms.MessageBox]::Show($Message, $Title, $ButtonType, $IconType)
    return $Result
}

function Wait-For-Acad {
    param($AcadApp)
    try {
        $maxWaitSeconds = 30
        $waitIntervalMs = 500
        $elapsedSeconds = 0
        # Wait for AutoCAD to reach a quiescent state (ready to accept commands)
        while ($AcadApp.GetAcadState().IsQuiescent -ne $true -and $elapsedSeconds -lt $maxWaitSeconds) {
            Start-Sleep -Milliseconds $waitIntervalMs
            $elapsedSeconds += ($waitIntervalMs / 1000)
        }
        if ($elapsedSeconds -ge $maxWaitSeconds) {
            Write-Log "WARNING: AutoCAD did not reach quiescent state within $maxWaitSeconds seconds."
        }
        Start-Sleep -Seconds 2 # Additional buffer time
    }
    catch {
        Write-Log "DEBUG: Warning - Could not check AutoCAD quiescent state. Forcing a 3-second wait."
        Start-Sleep -Seconds 3
    }
}

function Initialize-AutoCAD {
    $maxRetries = 3
    $retryCount = 0
    $Acad = $null
    $script:AcadWasStarted = $false

    while ($null -eq $Acad -and $retryCount -lt $maxRetries) {
        try {
            Write-Log "DEBUG: Attempting to connect to AutoCAD (Attempt $($retryCount + 1) of $maxRetries)..."
            $Acad = [System.Runtime.InteropServices.Marshal]::GetActiveObject("AutoCAD.Application")
            Write-Log "DEBUG: Successfully connected to running AutoCAD instance."
            $script:AcadWasStarted = $false
        }
        catch {
            try {
                Write-Log "DEBUG: No running AutoCAD instance found. Starting a new one..."
                $Acad = New-Object -ComObject "AutoCAD.Application"
                Write-Log "DEBUG: Started a new AutoCAD instance."
                $script:AcadWasStarted = $true
            }
            catch {
                Write-Log "ERROR: Failed to start or connect to AutoCAD: $($_.Exception.Message)"
                $retryCount++
                if ($retryCount -lt $maxRetries) {
                    Write-Log "DEBUG: Retrying in 5 seconds..."
                    Start-Sleep -Seconds 5
                }
                continue
            }
        }
    }

    if ($null -eq $Acad) {
        Write-Log "ERROR: Could not start or connect to AutoCAD after $maxRetries attempts. Please ensure AutoCAD is installed."
        exit 1
    }

    # --- FIX: Ensure AutoCAD is visible and wait for it to be ready ---
    $Acad.Visible = $true
    Wait-For-Acad -AcadApp $Acad
    
    # If a new instance was started, it often opens a default "Drawing1.dwg". Close it to prevent conflicts.
    if ($script:AcadWasStarted) {
        Write-Log "DEBUG: New instance started. Checking for and closing default drawing."
        try {
            # Check for Drawing1.dwg or Drawing2.dwg (sometimes it's the second one)
            $defaultDoc = $Acad.Documents | Where-Object { $_.Name -match "^Drawing\d+\.dwg$" } | Select-Object -First 1
            if ($defaultDoc -ne $null) {
                $defaultDoc.Close($false) # Close without saving
                Wait-For-Acad -AcadApp $Acad
                Write-Log "DEBUG: Closed default drawing: $($defaultDoc.Name)."
            }
        }
        catch {
            Write-Log "DEBUG: Could not close default drawing: $($_.Exception.Message)"
        }
    }
    # --- END FIX ---

    return $Acad
}

try {
    Write-Log "DEBUG: Script started. Initializing..."
    Write-Log "DEBUG: Received Titleblock Path: $TitleblockPath"
    Write-Log "DEBUG: Received CAD DWG Paths Count: $($CadDwgPaths.Count)"
    Write-Log "DEBUG: Received Titleblock Size: $TitleblockSize"
    Write-Log "DEBUG: Received Output Folder: $OutputFolder"

    # Initialize AutoCAD with retry logic
    $Acad = Initialize-AutoCAD
    # $Acad.Visible = $true # Removed, now in Initialize-AutoCAD

    # --- 1. Process Titleblock ---
    if (-not [string]::IsNullOrEmpty($TitleblockPath)) {
        Write-Log "INFO: Starting Titleblock processing."
        $Doc = $null
        try {
            Write-Log "DEBUG: Opening Titleblock: $(Split-Path $TitleblockPath -Leaf)"
            $Doc = $Acad.Documents.Open($TitleblockPath)
            $Doc.Activate()
            Wait-For-Acad -AcadApp $Acad
            
            Write-Log "DEBUG: Running CLEANTBLK command..."
            $Doc.SendCommand("._CLEANTBLK `n")
            Wait-For-Acad -AcadApp $Acad
            Write-Log "DEBUG: CLEANTBLK command finished. Waiting for user input."

            $PromptMessage = @"
Cleaning of titleblock '$(Split-Path $TitleblockPath -Leaf)' is complete.
Please review the drawing in AutoCAD.
- Click YES to SAVE the changes and continue.
- Click NO to UNDO the changes and handle it manually.
- Click CANCEL to LEAVE the changes but fix it manually.
"@
            $UserChoice = Show-UserPrompt -Message $PromptMessage

            if ($UserChoice -eq "Yes") {
                Write-Log "DEBUG: User chose YES. Saving and closing titleblock."
                $Doc.Save()
                $Doc.Close()
                Wait-For-Acad -AcadApp $Acad
            }
            elseif ($UserChoice -eq "No") {
                Write-Log "DEBUG: User chose NO. Undoing changes."
                $Doc.SendCommand("._UNDO 1 `n")
                Write-Log "INFO: Please close the titleblock manually in AutoCAD to continue."
                # Wait for user to close the drawing with timeout
                $maxWaitSeconds = 300
                $elapsedSeconds = 0
                while ($Acad.Documents.Count -gt 0 -and $elapsedSeconds -lt $maxWaitSeconds) {
                    Write-Log "DEBUG: Waiting for user to close the titleblock drawing..."
                    Start-Sleep -Seconds 3
                    $elapsedSeconds += 3
                }
                if ($Acad.Documents.Count -gt 0) {
                    Write-Log "ERROR: Timeout waiting for user to close the titleblock drawing."
                    exit 1
                }
            }
            else {
                # Cancel
                Write-Log "DEBUG: User chose CANCEL. Leaving changes."
                Write-Log "INFO: Please close the titleblock manually in AutoCAD to continue."
                $maxWaitSeconds = 300
                $elapsedSeconds = 0
                while ($Acad.Documents.Count -gt 0 -and $elapsedSeconds -lt $maxWaitSeconds) {
                    Write-Log "DEBUG: Waiting for user to close the titleblock drawing..."
                    Start-Sleep -Seconds 3
                    $elapsedSeconds += 3
                }
                if ($Acad.Documents.Count -gt 0) {
                    Write-Log "ERROR: Timeout waiting for user to close the titleblock drawing."
                    exit 1
                }
            }
        }
        catch {
            Write-Log "ERROR: An error occurred while processing the titleblock: $($_.Exception.Message)"
            if ($null -ne $Doc) { $Doc.Close($false) }
            exit 1
        }
    }
    else {
        Write-Log "ERROR: No titleblock path provided."
        exit 1
    }

    # --- 2. Process CAD DWGs ---
    Write-Log "INFO: Titleblock processing complete. Moving to CAD drawings."
    
    if ($null -eq $CadDwgPaths -or $CadDwgPaths.Count -eq 0) {
        Write-Log "ERROR: No CAD DWG paths were received by the script. Aborting."
        exit 1
    }

    Write-Log "INFO: Found $($CadDwgPaths.Count) CAD DWG(s) to process."
    $fileCounter = 0

    foreach ($DwgPath in $CadDwgPaths) {
        $fileCounter++
        Write-Log "INFO: Processing file $fileCounter of $($CadDwgPaths.Count): $(Split-Path $DwgPath -Leaf)"
        $Doc = $null
        try {
            # Ensure no documents are open before proceeding with timeout
            $maxWaitSeconds = 300
            $elapsedSeconds = 0
            while ($Acad.Documents.Count -gt 0 -and $elapsedSeconds -lt $maxWaitSeconds) {
                Write-Log "DEBUG: Waiting for user to close any open drawings..."
                Start-Sleep -Seconds 3
                $elapsedSeconds += 3
            }
            if ($Acad.Documents.Count -gt 0) {
                Write-Log "ERROR: Timeout waiting for user to close open drawings."
                exit 1
            }

            Write-Log "DEBUG: Opening DWG: $(Split-Path $DwgPath -Leaf)"
            $Doc = $Acad.Documents.Open($DwgPath)
            $Doc.Activate()
            Wait-For-Acad -AcadApp $Acad

            Write-Log "DEBUG: Running CLEANCAD command..."
            $Doc.SendCommand("._CLEANCAD `n")
            Wait-For-Acad -AcadApp $Acad
            Write-Log "DEBUG: CLEANCAD command finished. Waiting for user input."

            $PromptMessage = @"
Cleaning of '$(Split-Path $DwgPath -Leaf)' is complete.
Please review the drawing in AutoCAD.
- Click YES to SAVE the changes (as a new file named after layouts) and continue.
- Click NO to UNDO the changes and handle it manually.
- Click CANCEL to LEAVE the changes but fix it manually.
"@
            $UserChoice = Show-UserPrompt -Message $PromptMessage

            if ($UserChoice -eq "Yes") {
                Write-Log "DEBUG: User chose YES. Getting layout names for new filename."
                $layoutNames = @()
                foreach ($layout in $Doc.Layouts) {
                    if ($layout.Name -ne "Model") { $layoutNames += $layout.Name }
                }

                if ($layoutNames.Count -gt 0) {
                    # The requirement is to save the file in "current folder/AutoCAD Clean DWGs"
                    # The output folder is already the correct local, structured, timestamped folder.
                    $newFileName = ($layoutNames -join " ") + ".dwg"
                    $newFilePath = Join-Path $OutputFolder $newFileName
                    Write-Log "DEBUG: Saving as '$newFilePath'"
                    $Doc.SaveAs($newFilePath)
                    Wait-For-Acad -AcadApp $Acad
                    $Doc.Close($false)
                }
                else {
                    Write-Log "DEBUG: No layouts found besides Model. Saving and closing original file."
                    # Note: This saves the *duplicated* file in the local structure.
                    $Doc.Save()
                    $Doc.Close()
                }
                Wait-For-Acad -AcadApp $Acad
            }
            elseif ($UserChoice -eq "No") {
                Write-Log "DEBUG: User chose NO. Undoing changes."
                $Doc.SendCommand("._UNDO 1 `n")
                Write-Log "INFO: Please close the DWG manually in AutoCAD to continue."
                $maxWaitSeconds = 300
                $elapsedSeconds = 0
                while ($Acad.Documents.Count -gt 0 -and $elapsedSeconds -lt $maxWaitSeconds) {
                    Write-Log "DEBUG: Waiting for user to close the DWG..."
                    Start-Sleep -Seconds 3
                    $elapsedSeconds += 3
                }
                if ($Acad.Documents.Count -gt 0) {
                    Write-Log "ERROR: Timeout waiting for user to close the DWG."
                    exit 1
                }
            }
            else {
                # Cancel
                Write-Log "DEBUG: User chose CANCEL. Leaving changes."
                Write-Log "INFO: Please close the DWG manually in AutoCAD to continue."
                $maxWaitSeconds = 300
                $elapsedSeconds = 0
                while ($Acad.Documents.Count -gt 0 -and $elapsedSeconds -lt $maxWaitSeconds) {
                    Write-Log "DEBUG: Waiting for user to close the DWG..."
                    Start-Sleep -Seconds 3
                    $elapsedSeconds += 3
                }
                if ($Acad.Documents.Count -gt 0) {
                    Write-Log "ERROR: Timeout waiting for user to close the DWG."
                    exit 1
                }
            }
        }
        catch {
            Write-Log "ERROR: An error occurred while processing '$DwgPath': $($_.Exception.Message)"
            if ($null -ne $Doc) { $Doc.Close($false) }
            exit 1
        }
    }

    # --- 3. Open Output Folder ---
    Write-Log "INFO: All files processed. Opening output folder."
    if (Test-Path $OutputFolder) {
        Invoke-Item $OutputFolder
    }
    else {
        Write-Log "ERROR: Output folder not found: $OutputFolder"
    }

    Write-Log "INFO: Script finished successfully."
}
catch {
    $ErrorMessage = $_.Exception.Message
    Write-Log "FATAL: A critical script error occurred: $ErrorMessage"
    exit 1
}
finally {
    if ($null -ne $Acad) {
        try {
            # Ensure all documents are closed
            foreach ($doc in $Acad.Documents) {
                $doc.Close($false)
            }
            # Only quit AutoCAD if we started it
            if ($script:AcadWasStarted) {
                $Acad.Quit()
            }
        }
        catch {
            Write-Log "DEBUG: Could not properly close AutoCAD: $($_.Exception.Message)"
        }
    }
}