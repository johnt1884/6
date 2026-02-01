@echo off
setlocal EnableExtensions EnableDelayedExpansion

:MENU
cls
echo ======================================
echo   SC Utilities
echo ======================================
echo.
echo   UPDATES
echo   ------------------
echo   1. Update scdate.txt (newest shortcut date)
echo   2. Update scdata.txt (shortcut listing)
echo   3. Generate scnew.txt (for Load SC New)
echo   4. Update selections.txt (for each project)
echo   5. Perform ALL updates (1-4)
echo.
echo   TOOLS
echo   ------------------
echo   6. Check Thumbnails
echo.
echo   7. Exit
echo.
set /p choices=Choose one or more options (e.g., 1 3 6):

for %%c in (%choices%) do (
    if "%%c"=="1" call :SCDATE
    if "%%c"=="2" call :SCDATA
    if "%%c"=="3" call :SCNEW
    if "%%c"=="4" call :SELECTIONS
    if "%%c"=="5" call :ALL_UPDATES
    if "%%c"=="6" call :CHECK_THUMBNAILS
    if "%%c"=="7" goto END
)

goto DONE


:ALL_UPDATES
call :SCDATE
call :SCDATA
call :SCNEW
call :SELECTIONS
goto :EOF

:SCDATE
call :DO_SCDATE
goto :EOF

:SCDATA
call :DO_SCDATA
goto :EOF

:SCNEW
call :DO_SCNEW
goto :EOF

:SELECTIONS
call :DO_SELECTIONS
goto :EOF

:CHECK_THUMBNAILS
echo.
echo Checking thumbnails...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0\check_thumbnails.ps1"
goto :EOF


:: --------------------------------------------------
:: Update scdate.txt (newest .lnk timestamp)
:: --------------------------------------------------
:DO_SCDATE
echo.
echo Updating scdate.txt files...

powershell -NoProfile -Command ^
    "$ws = New-Object -ComObject WScript.Shell; " ^
    "$root = $PWD.Path; " ^
    "$rootSc = Join-Path $root 'sc'; " ^
    "$cached = @(); " ^
    "if (Test-Path $rootSc) { " ^
    "  Get-ChildItem $rootSc -Filter *.lnk | ForEach-Object { " ^
    "    try { " ^
    "      $t = $ws.CreateShortcut($_.FullName).TargetPath; " ^
    "      if ($t) { $cached += @{ Target=$t; Date=$_.LastWriteTime.ToUniversalTime() } } " ^
    "    } catch {} " ^
    "  } " ^
    "}; " ^
    "$targetDirs = New-Object System.Collections.Generic.HashSet[string]; " ^
    "$null = $targetDirs.Add($root); " ^
    "Get-ChildItem -Path $root -Filter sc -Directory -Recurse | ForEach-Object { " ^
    "  $null = $targetDirs.Add((Split-Path $_.FullName -Parent)) " ^
    "}; " ^
    "Get-ChildItem -Path $root -Directory | Where-Object { $_.Name -notmatch '^(sc|landscape|landscape rotate|edit|thumbnails|edit thumbnails)$' } | ForEach-Object { " ^
    "  $null = $targetDirs.Add($_.FullName) " ^
    "}; " ^
    "foreach ($dir in $targetDirs) { " ^
    "  $out = Join-Path $dir 'scdate.txt'; " ^
    "  $newest = [DateTime]::MinValue; " ^
    "  $pSc = Join-Path $dir 'sc'; " ^
    "  if (Test-Path $pSc) { " ^
    "    $lnk = Get-ChildItem $pSc -Filter *.lnk | Sort-Object LastWriteTime -Descending | Select-Object -First 1; " ^
    "    if ($lnk) { $newest = $lnk.LastWriteTime.ToUniversalTime() } " ^
    "  }; " ^
    "  if ($dir -ne $root) { " ^
    "    foreach ($c in $cached) { " ^
    "      if ($c.Target.StartsWith($dir + [IO.Path]::DirectorySeparatorChar) -or $c.Target -eq $dir) { " ^
    "        if ($c.Date -gt $newest) { $newest = $c.Date } " ^
    "      } " ^
    "    } " ^
    "  }; " ^
    "  if ($newest -gt [DateTime]::MinValue) { " ^
    "    $write = $true; " ^
    "    if (Test-Path $out) { " ^
    "      try { " ^
    "        $content = Get-Content $out -Raw; " ^
    "        if ($content.StartsWith('dummy:')) { " ^
    "          $dDate = [DateTimeOffset]::Parse($content.Substring(6).Trim()).UtcDateTime; " ^
    "          if ($newest -le $dDate) { $write = $false } " ^
    "        } " ^
    "      } catch {} " ^
    "    }; " ^
    "    if ($write) { " ^
    "      $isoDate = $newest.ToString('yyyy-MM-ddTHH:mm:ss.fffZ'); " ^
    "      Set-Content -Path $out -Value $isoDate -Encoding ASCII " ^
    "    } " ^
    "  } " ^
    "}"

exit /b


:: --------------------------------------------------
:: Update scdata.txt
:: --------------------------------------------------
:DO_SCDATA
echo.
echo Updating scdata.txt files...

:: Recursive sc folders (UNCHANGED behavior)
for /d /r %%D in (sc) do (
    if exist "%%D\*.lnk" (
        dir "%%D\*.lnk" /b > "%%D\..\scdata.txt"
    )
)

:: Top-level ".\sc" (grouped target output — FIXED)
if exist "%CD%\sc\*.lnk" (
    powershell -NoProfile -Command ^
        "$out = Join-Path (Get-Location) 'scdata.txt';" ^
        "Remove-Item $out -ErrorAction SilentlyContinue;" ^
        "$ws = New-Object -ComObject WScript.Shell;" ^
        "$groups = @{};" ^
        "Get-ChildItem '.\sc' -Filter *.lnk | ForEach-Object {" ^
        "  $t = $ws.CreateShortcut($_.FullName).TargetPath;" ^
        "  if ($t) {" ^
        "    $folder = Split-Path (Split-Path $t -Parent) -Leaf;" ^
        "    $file = Split-Path $t -Leaf;" ^
        "    if (-not $groups.ContainsKey($folder)) { $groups[$folder] = @() };" ^
        "    $groups[$folder] += $file;" ^
        "  }" ^
        "};" ^
        "$groups.Keys | Sort-Object | ForEach-Object {" ^
        "  Add-Content $out ('\"' + $_ + '\"');" ^
        "  $groups[$_] | Sort-Object | ForEach-Object { Add-Content $out $_ };" ^
        "  Add-Content $out '';" ^
        "}"
)

exit /b


:: --------------------------------------------------
:: Generate scnew.txt (for Load SC New)
:: --------------------------------------------------
:DO_SCNEW
echo.
echo Generating scnew.txt...

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0\generate_scnew.ps1"

exit /b


:: --------------------------------------------------
:: Update selections.txt for each project subfolder
:: --------------------------------------------------
:DO_SELECTIONS
echo.
echo Updating selections.txt files...

REM Process only immediate subdirectories of the current directory
for /d %%D in (*) do (
    REM Get just the folder name for comparison
    set "folderName=%%~nxD"

    REM Create a lowercase version for case-insensitive comparison
    set "lowerFolderName=!folderName!"
    for %%C in (A B C D E F G H I J K L M N O P Q R S T U V W X Y Z) do (
        set "lowerFolderName=!lowerFolderName:%%C=%%c!"
    )

    REM Check if the folder is one of the special ones to be skipped
    set "isSpecial=0"
    if "!lowerFolderName!"=="sc" set "isSpecial=1"
    if "!lowerFolderName!"=="landscape" set "isSpecial=1"
    if "!lowerFolderName!"=="landscape rotate" set "isSpecial=1"
    if "!lowerFolderName!"=="edit" set "isSpecial=1"
    if "!lowerFolderName!"=="thumbnails" set "isSpecial=1"
    if "!lowerFolderName!"=="edit thumbnails" set "isSpecial=1"

    if "!isSpecial!"=="0" (
        echo Processing folder: %%D
        set "OUT=%%D\selections.txt"
        > "!OUT!" type nul

        for %%F in ("sc" "Landscape" "Landscape Rotate" "Edit") do (
            echo # %%~F>> "!OUT!"
            if exist "%%D\%%~F\" (
                for /f "delims=" %%A in ('dir /b /a:-d "%%D\%%~F" 2^>nul') do (
                    echo %%A>> "!OUT!"
                )
            )
            echo.>> "!OUT!"
        )
    )
)

exit /b



:DONE
echo.
echo Done.
pause
goto MENU


:END
endlocal
exit /b
