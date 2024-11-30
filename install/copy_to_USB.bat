rem *** REWRITE THE REPORT GENERATOR EXE AND FOLDERS FROM DESKTOP PC TO A PREVIOUSLY configured USB DRIVE ***

@echo off
setlocal enabledelayedexpansion

rem Detect the USB drive, which is the target directory
set "usbDrive="
for %%i in (D E F G H I J K L M N O P Q R S T U V W X Y Z) do (
    set "drive=%%i:"
    if exist "!drive!\TIMC_APP.exe" (
        set "usbDrive=!drive!"
        goto :foundDrive
    )
)

:foundDrive
if not defined usbDrive (
    echo previously configured USB drive not found.
    goto :end
)

echo Deleting all files on the USB drive...
del /Q %usbDrive%*.*

rem Set the source directory on the PC desktop
set "sourceDir=%USERPROFILE%\Desktop\TIMC_APP\"

echo Copying files from desktop to USB drive...
xcopy /E /Y "%sourceDir%*" %usbDrive%

echo Files copied to %usbDrive% from %SourceDir%
echo Copy completed


:end

endlocal

rem Pause to keep the command prompt window open
pause