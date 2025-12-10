@echo off
setlocal

:: Check for admin rights and self-elevate if needed
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Requesting administrator privileges...
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

:: Configuration
set "DLL_NAME=WinRARShellExtQuickExtract.dll"
set "INSTALL_DIR=%ProgramFiles%\WinRAR"
set "SOURCE_DIR=%~dp0"

echo.
echo ========================================
echo  WinRAR Quick Extract - Installer
echo ========================================
echo.

:: Check if WinRAR folder exists
if not exist "%INSTALL_DIR%" (
    echo ERROR: WinRAR installation not found at %INSTALL_DIR%
    pause
    exit /b 1
)

:: Check if DLL exists in same folder as this script
if not exist "%SOURCE_DIR%%DLL_NAME%" (
    echo ERROR: %DLL_NAME% not found!
    echo Please run build.bat first.
    pause
    exit /b 1
)

:: Stop Explorer to release any locks on the DLL
echo Stopping Explorer...
taskkill /f /im explorer.exe >nul 2>&1
timeout /t 2 /nobreak >nul

:: Copy DLL to WinRAR directory
echo Copying %DLL_NAME% to %INSTALL_DIR%...
copy /y "%SOURCE_DIR%%DLL_NAME%" "%INSTALL_DIR%\%DLL_NAME%" >nul
if %errorlevel% neq 0 (
    echo ERROR: Failed to copy DLL!
    start explorer.exe
    pause
    exit /b 1
)

:: Register the DLL
echo Registering shell extension...
regsvr32 /s "%INSTALL_DIR%\%DLL_NAME%"
if %errorlevel% neq 0 (
    echo ERROR: Failed to register DLL!
    start explorer.exe
    pause
    exit /b 1
)

:: Restart Explorer
echo Restarting Explorer...
start explorer.exe

echo.
echo ========================================
echo  Installation complete!
echo.
echo  The "Extract to folder" option will now
echo  appear in the context menu for archives.
echo ========================================
echo.
pause
