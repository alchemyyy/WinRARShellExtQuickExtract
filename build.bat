@echo off
setlocal

:: Find Visual Studio
set "VSDIR=C:\Program Files\Microsoft Visual Studio\18\Enterprise"
if not exist "%VSDIR%" set "VSDIR=C:\Program Files\Microsoft Visual Studio\2022\Community"
if not exist "%VSDIR%" set "VSDIR=C:\Program Files\Microsoft Visual Studio\2022\Professional"
if not exist "%VSDIR%" set "VSDIR=C:\Program Files\Microsoft Visual Studio\2022\Enterprise"

if not exist "%VSDIR%\VC\Auxiliary\Build\vcvars64.bat" (
    echo ERROR: Visual Studio not found!
    pause
    exit /b 1
)

:: Setup
set "PROJECT_DIR=%~dp0"
set "BUILD_DIR=%PROJECT_DIR%build"

if not exist "%BUILD_DIR%" mkdir "%BUILD_DIR%"

echo Building WinRARShellExtQuickExtract.dll...

call "%VSDIR%\VC\Auxiliary\Build\vcvars64.bat" >nul

cl /nologo /O2 /W3 /LD /DUNICODE /D_UNICODE ^
   "%PROJECT_DIR%main.c" ^
   /Fo:"%BUILD_DIR%\\" ^
   /Fe:"%BUILD_DIR%\WinRARShellExtQuickExtract.dll" ^
   /link /DEF:"%PROJECT_DIR%WinRARShellExtQuickExtract.def" ^
   ole32.lib shell32.lib advapi32.lib user32.lib shlwapi.lib comctl32.lib gdi32.lib

if %errorlevel% neq 0 (
    echo Build failed!
    pause
    exit /b 1
)

:: Copy install/uninstall scripts to build folder
echo Copying install scripts to build folder...
copy /y "%PROJECT_DIR%install.bat" "%BUILD_DIR%\install.bat" >nul
copy /y "%PROJECT_DIR%uninstall.bat" "%BUILD_DIR%\uninstall.bat" >nul

echo.
echo ========================================
echo Build successful!
echo.
echo Output folder: %BUILD_DIR%
echo   - WinRARShellExtQuickExtract.dll
echo   - install.bat
echo   - uninstall.bat
echo ========================================

pause
