param(
    [string]$TitleblockPath,
    [string]$Disciplines,
    [string]$TitleblockSize,
    [string]$OutputFolder,
    [string]$ProjectRoot
)

# --- DEBUG LOGGING ---
Write-Host "PROGRESS: ===== CLEAN DWGS SCRIPT STARTED ====="
Write-Host "PROGRESS: Auto-including titleblock parent folder"
Write-Host "PROGRESS: TitleblockPath: '$TitleblockPath'"
Write-Host "PROGRESS: Folders to process: '$Disciplines'"
Write-Host "PROGRESS: OutputFolder: '$OutputFolder'"

# ... rest of the PowerShell script ...

# Split disciplines string into array
$DisciplineList = $Disciplines -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }

Write-Host "PROGRESS: Processing folders: $($DisciplineList -join ' | ')"

# WINDOW MANAGEMENT & HELPER FUNCTIONS...

# --- MAIN SCRIPT ---
try {
    Write-Log "=== SCRIPT EXECUTION BEGINS ==="
    
    if ($null -eq $DisciplineList -or $DisciplineList.Count -eq 0) {
        throw "No folders provided."
    }
    
    $Acad = Initialize-AutoCAD

    # --- 1. Process Titleblock ---
    if (-not [string]::IsNullOrEmpty($TitleblockPath)) {
        Write-Log "Processing Titleblock: $([System.IO.Path]::GetFileName($TitleblockPath))"
        $Doc = $null
        $tbName = [System.IO.Path]::GetFileName($TitleblockPath)
        
        try {
            $Doc = $Acad.Documents.Open($TitleblockPath)
            $Doc.Activate()
            Bring-Acad-To-Front -AcadApp $Acad
            Wait-For-Acad -AcadApp $Acad
            
            $Doc.SendCommand("._CLEANTBLK `n")
            Wait-For-Acad -AcadApp $Acad

            $PromptMessage = "Cleaning of titleblock '$tbName' complete.`n`nYES: Save and continue`nNO: Undo`nCANCEL: Leave changes"
            $UserChoice = Show-UserPrompt -Message $PromptMessage

            if ($UserChoice -eq "Yes") {
                $Doc.Save()
                $Doc.Close()
            }
            else {
                if ($UserChoice -eq "No") { $Doc.SendCommand("._UNDO 1 `n") }
                # Wait for manual close...
            }
            Wait-For-Acad -AcadApp $Acad
        }
        catch {
            Write-Log "ERROR: Titleblock failed: $($_.Exception.Message)"
            if ($null -ne $Doc) { try { $Doc.Close($false) } catch {} }
            throw
        }
    }

    # --- 2. Discover and Process DWG Files ---
    Write-Log "Discovering DWG files in folders: $($DisciplineList -join ' | ')"
    
    $AllDwgFiles = @()
    $OutputRootPath = [System.IO.Path]::GetFullPath($OutputFolder)
    
    foreach ($folder in $DisciplineList) {
        $folderPath = Join-Path $OutputRootPath $folder
        if (Test-Path $folderPath) {
            Write-Log "Searching in: $folderPath"
            $dwgFiles = Get-ChildItem -Path $folderPath -Filter "*.dwg" -File -Recurse | 
            Where-Object { $_.FullName -ne $TitleblockPath } |
            Select-Object -ExpandProperty FullName
            Write-Log "Found $($dwgFiles.Count) DWG files in $folder"
            $AllDwgFiles += $dwgFiles
        }
        else {
            Write-Log "WARNING: Folder not found: $folderPath"
        }
    }
    
    Write-Log "Total DWG files to process: $($AllDwgFiles.Count)"
    
    if ($AllDwgFiles.Count -eq 0) {
        throw "No DWG files found in selected folders."
    }

    # Process each DWG file
    $fileCounter = 0
    foreach ($DwgPath in $AllDwgFiles) {
        $fileCounter++
        $docName = [System.IO.Path]::GetFileName($DwgPath)
        Write-Log "=== FILE $fileCounter of $($AllDwgFiles.Count): $docName ==="
        $Doc = $null
        
        try {
            # Verify AutoCAD is responsive
            try { $test = $Acad.Visible } catch {
                Write-Log "WARNING: AutoCAD COM object died. Reinitializing..."
                try { $Acad.Quit() } catch {}
                $Acad = Initialize-AutoCAD
            }
            
            # Check if already open
            $Doc = $Acad.Documents | Where-Object { $_.Name -eq $docName } | Select-Object -First 1
            
            if ($null -eq $Doc) {
                if (-not (Test-Path $DwgPath)) {
                    throw "File not found: $DwgPath"
                }
                $Doc = $Acad.Documents.Open($DwgPath)
            }
            
            $Doc.Activate()
            Bring-Acad-To-Front -AcadApp $Acad
            Wait-For-Acad -AcadApp $Acad
            
            $Doc.SendCommand("._CLEANCAD `n")
            Wait-For-Acad -AcadApp $Acad

            $PromptMessage = "Cleaning of '$docName' complete.`n`nYES: Save (by layouts) and continue`nNO: Undo`nCANCEL: Leave changes"
            $UserChoice = Show-UserPrompt -Message $PromptMessage

            if ($UserChoice -eq "Yes") {
                $layoutNames = @()
                foreach ($layout in $Doc.Layouts) {
                    if ($layout.Name -ne "Model") { $layoutNames += $layout.Name }
                }

                if ($layoutNames.Count -gt 0) {
                    $newFileName = ($layoutNames -join " ") + ".dwg"
                    $newFilePath = Join-Path $OutputFolder $newFileName
                    $Doc.SaveAs($newFilePath)
                    $Doc.Close($false)
                }
                else {
                    $Doc.Save()
                    $Doc.Close()
                }
            }
            else {
                if ($UserChoice -eq "No") { $Doc.SendCommand("._UNDO 1 `n") }
                # Wait for manual close...
            }
            
            # Ensure cleanup
            if ($null -ne $Doc) {
                try { $test = $Doc.Name; $Doc.Close($false) } catch {}
            }
            
            Wait-For-Acad -AcadApp $Acad
            Start-Sleep -Seconds 3
            
        }
        catch {
            Write-Log "ERROR: Failed to process '$docName': $($_.Exception.Message)"
            if ($null -ne $Doc) { try { $Doc.Close($false) } catch {} }
            Write-Log "Continuing to next file..."
            Wait-For-Acad -AcadApp $Acad
            Start-Sleep -Seconds 3
            continue
        }
    }
    
    Write-Log "INFO: All CAD files processed."
}
catch {
    $ErrorMessage = $_.Exception.Message
    Write-Log "FATAL: Critical error: $ErrorMessage"
}
finally {
    if ($null -ne $Acad) {
        try {
            foreach ($doc in $Acad.Documents) { try { $doc.Close($false) } catch {} }
            if ($script:AcadWasStarted) { $Acad.Quit() }
        }
        catch {}
    }
}

# --- Open Output Folder ---
if (Test-Path $OutputFolder) {
    Invoke-Item $OutputFolder
}
else {
    Write-Log "ERROR: Output folder not found: $OutputFolder"
}

Write-Log "=== SCRIPT FINISHED ==="