$extractDir = Join-Path $env:TEMP "acad_extract"
Write-Host "Extract Directory: $extractDir"

if (Test-Path $extractDir) {
    $files = Get-ChildItem $extractDir
    foreach ($f in $files) {
        Write-Host "File: $($f.Name) - Size: $($f.Length)"
        if ($f.Name -like "*.txt") {
            Write-Host "--- CONTENT START ---"
            Get-Content $f.FullName | Select-Object -First 50
            Write-Host "--- CONTENT END ---"
        }
    }
}
else {
    Write-Host "Directory not found."
}
