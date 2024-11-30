rem *** BATCH FILE TO INSTALL TIMC REPORT GENERATOR EXECUTABLE ON A DESKTOP PC FROM A USB DRIVE ***

@echo off
setlocal enabledelayedexpansion

rem Detect the USB drive which is the source directory
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
    echo USB drive not found.
    goto :end
)

rem Check if correct version of Visual C++ Redistributable is installed on the PC:

rem Define the registry key for Visual C++ Redistributable x64
set "vcRedistKey=HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64"

rem Check if the registry key exists
reg query "%vcRedistKey%" >nul 2>&1

if %errorlevel% neq 0 (
    echo Visual C++ x64 is not installed. Installing...

    rem Check if the installer file exists on the USB drive
    set "installerPath=%usbDrive%\VC_redist.x64.exe"
    if exist "%installerPath%" (
        rem Run the installer
         start /wait "" "%installerPath%" /install /passive /norestart

        rem Check again after installation
        reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64" >nul 2>&1
        if %errorlevel% equ 0 (
            echo Installation successful.
        ) else (
            echo Installation failed.
        )
    ) else (
        echo Installer not found on the USB drive.
    )
) else (
    echo Visual C++ x64 is already installed.
)

rem Set the destination directory on the PC desktop
set "destinationDir=%USERPROFILE%\Desktop\TIMC_APP"

rem Create the destination directory if it doesn't exist
if not exist "%destinationDir%" mkdir "%destinationDir%"

rem Change the current directory to the desktop
cd /d "%USERPROFILE%\Desktop"

rem Copy files and subdirectories from USB drive to destination
xcopy /s /i /e "%usbDrive%\*" "%destinationDir%"

echo Files copied from %usbDrive% to %destinationDir%
echo Copy completed.

rem create a shortcut

set "TargetPath=%USERPROFILE%\Desktop\TIMC_APP\TIMC_APP.exe"
set "ShortcutPath=%USERPROFILE%\Desktop\TIMC Report Generator.lnk"
powershell -Command "$WshShell = New-Object -ComObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%ShortcutPath%'); $Shortcut.TargetPath = '%TargetPath%'; $Shortcut.Save()"
echo Shortcut created

endlocal

:end

rem Pause to keep the command prompt window open
pause 
