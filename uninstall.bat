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

echo.
echo ========================================
echo  WinRAR Quick Extract - Uninstaller
echo ========================================
echo.

:: Stop Explorer to release locks
echo Stopping Explorer...
taskkill /f /im explorer.exe >nul 2>&1
timeout /t 2 /nobreak >nul

:: Unregister the DLL
echo Unregistering shell extension...
if exist "%INSTALL_DIR%\%DLL_NAME%" (
    regsvr32 /s /u "%INSTALL_DIR%\%DLL_NAME%"
    del /f "%INSTALL_DIR%\%DLL_NAME%"
)

:: Restart Explorer
echo Restarting Explorer...
start explorer.exe

echo.
echo ========================================
echo  Uninstallation complete!
echo ========================================
echo.
pause
