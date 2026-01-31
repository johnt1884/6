# generate_scnew.ps1
# Root = where the batch file is run from
$root = Get-Location
$ws = New-Object -ComObject WScript.Shell

# Root sc folder and shortcuts
$rootSc = Join-Path $root 'sc'
$rootLinks = @() # Initialize as empty array
if (Test-Path $rootSc) {
    # Cache root shortcuts if they exist
    $rootLinks = Get-ChildItem "$rootSc\*.lnk" -ErrorAction SilentlyContinue
}

# Process ONLY top-level subfolders (excluding the root 'sc' folder itself)
Get-ChildItem $root -Directory | Where-Object { $_.Name -ne 'sc' } | ForEach-Object {

    $proj = $_
    $projSc = Join-Path $proj.FullName 'sc'
    $scnewFile = Join-Path $proj.FullName 'scnew.txt'
    
    # Example 4: Subfolder has no 'sc' directory or no shortcuts within it.
    # The scnew.txt file should be deleted if it exists.
    if (-not (Test-Path $projSc) -or -not (Get-ChildItem "$projSc\*.lnk" -ErrorAction SilentlyContinue)) {
        if (Test-Path $scnewFile) {
            Remove-Item $scnewFile
        }
        return # Equivalent to 'continue'
    }

    # Find which of the root shortcuts point to the current project subfolder
    $matchingRootLinks = @()
    if ($rootLinks.Count -gt 0) {
        $matchingRootLinks = foreach ($lnk in $rootLinks) {
            $target = $ws.CreateShortcut($lnk.FullName).TargetPath
            if ($target -and $target.StartsWith($proj.FullName, [StringComparison]::OrdinalIgnoreCase)) {
                $lnk # Add the matching link object to the collection
            }
        }
    }

    $newLinks = @()

    # Example 1: Root sc has no shortcuts to this project.
    # All local shortcuts are considered "new".
    if ($matchingRootLinks.Count -eq 0) {
        $newLinks = Get-ChildItem "$projSc\*.lnk" | Sort-Object LastWriteTime
    } else {
        # Example 2 & 3: Root sc has shortcuts to this project.
        # Compare dates to find newer local shortcuts.
        $cutoff = ($matchingRootLinks |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1).LastWriteTime

        $newLinks = Get-ChildItem "$projSc\*.lnk" |
            Where-Object { $_.LastWriteTime -gt $cutoff } |
            Sort-Object LastWriteTime
    }

    # If new links were found, create/overwrite scnew.txt.
    if ($newLinks.Count -gt 0) {
        $newLinks |
            Select-Object -ExpandProperty Name |
            Set-Content -Path $scnewFile -Encoding ASCII
    } else {
        # Otherwise (Example 3), delete scnew.txt if it exists.
        if (Test-Path $scnewFile) {
            Remove-Item $scnewFile
        }
    }
}
