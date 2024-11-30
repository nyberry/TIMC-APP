@echo off
setlocal enabledelayedexpansion

rem Define source directory
set "source=%~dp0"

rem Define destination drive letters to check
set "driveletters=D E F G"

rem Loop through each drive letter
for %%d in (%driveletters%) do (
    rem Check if the drive is a removable drive
    fsutil fsinfo drivetype %%d: | find "Removable" >nul
    if not errorlevel 1 (
        rem If removable drive found, copy files and exit loop
        set "destination=%%d:"
        goto :copyfiles
    )
)

:copyfiles
rem Check if destination is set
if defined destination (
    echo Copying files to %destination%
    xcopy /s /e "%source%" "%destination%"
    echo Files copied successfully.
) else (
    echo No removable drive found.
)
pause