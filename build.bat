@echo off
:: ============================================================
::  OpenSheet ET — Windows Build Script
::  Requirements:
::    Qt 6.5+  (MSVC 2019/2022 x64)
::    CMake 3.22+
::    Visual Studio 2019 or 2022 (with C++ Desktop workload)
::    NSIS 3.x (for installer, optional)
::
::  Usage:
::    build.bat                          (uses defaults below)
::    build.bat Release "C:\Qt\6.5.3\msvc2022_64"
:: ============================================================

setlocal EnableDelayedExpansion

:: ── Configuration ────────────────────────────────────────────────────────────
set BUILD_TYPE=%1
set QT_DIR=%2

if "%BUILD_TYPE%"=="" set BUILD_TYPE=Release
if "%QT_DIR%"=="" (
    :: Auto-detect common Qt paths
    for %%P in (
        "C:\Qt\6.8.0\msvc2022_64"
        "C:\Qt\6.7.0\msvc2022_64"
        "C:\Qt\6.6.0\msvc2022_64"
        "C:\Qt\6.5.3\msvc2022_64"
        "C:\Qt\6.5.3\msvc2019_64"
        "C:\Qt\6.5.0\msvc2019_64"
        "D:\Qt\6.5.3\msvc2022_64"
    ) do (
        if exist %%P\bin\qmake.exe (
            set QT_DIR=%%~P
            goto :found_qt
        )
    )
    echo ERROR: Qt not found. Pass Qt path as second argument:
    echo   build.bat Release "C:\Qt\6.5.3\msvc2022_64"
    exit /b 1
)
:found_qt

set BUILD_DIR=build-windows
set DEPLOY_DIR=deploy
set APP_NAME=OpenSheetET
set APP_EXE=%APP_NAME%.exe

echo.
echo ╔══════════════════════════════════════════════════════════╗
echo ║          OpenSheet ET — Windows Build System            ║
echo ╠══════════════════════════════════════════════════════════╣
echo ║  Build type : %BUILD_TYPE%
echo ║  Qt path    : %QT_DIR%
echo ║  Build dir  : %BUILD_DIR%
echo ║  Deploy dir : %DEPLOY_DIR%
echo ╚══════════════════════════════════════════════════════════╝
echo.

:: ── Step 1: Configure ────────────────────────────────────────────────────────
echo [1/5] Configuring CMake...
cmake -B "%BUILD_DIR%" ^
      -DCMAKE_BUILD_TYPE=%BUILD_TYPE% ^
      -DCMAKE_PREFIX_PATH="%QT_DIR%" ^
      -G "Visual Studio 17 2022" ^
      -A x64
if errorlevel 1 (
    echo.
    echo Trying Visual Studio 16 2019...
    cmake -B "%BUILD_DIR%" ^
          -DCMAKE_BUILD_TYPE=%BUILD_TYPE% ^
          -DCMAKE_PREFIX_PATH="%QT_DIR%" ^
          -G "Visual Studio 16 2019" ^
          -A x64
    if errorlevel 1 goto :error
)
echo   OK: CMake configured.

:: ── Step 2: Build ────────────────────────────────────────────────────────────
echo.
echo [2/5] Building %APP_NAME%...
cmake --build "%BUILD_DIR%" --config %BUILD_TYPE% --parallel
if errorlevel 1 goto :error
echo   OK: Build complete.

:: ── Step 3: Prepare deploy directory ─────────────────────────────────────────
echo.
echo [3/5] Preparing deployment...
if exist "%DEPLOY_DIR%" rmdir /s /q "%DEPLOY_DIR%"
mkdir "%DEPLOY_DIR%"
copy /y "%BUILD_DIR%\%BUILD_TYPE%\%APP_EXE%" "%DEPLOY_DIR%\%APP_EXE%"
if errorlevel 1 goto :error
echo   OK: Executable copied.

:: ── Step 4: windeployqt ───────────────────────────────────────────────────────
echo.
echo [4/5] Running windeployqt...
"%QT_DIR%\bin\windeployqt.exe" ^
    --dir "%DEPLOY_DIR%" ^
    --no-translations ^
    --no-system-d3d-compiler ^
    --no-opengl-sw ^
    --release ^
    "%DEPLOY_DIR%\%APP_EXE%"
if errorlevel 1 (
    echo   WARNING: windeployqt failed, continuing anyway...
) else (
    echo   OK: Qt runtime deployed.
)

:: ── Step 5: NSIS Installer ────────────────────────────────────────────────────
echo.
echo [5/5] Building installer...
where makensis >nul 2>&1
if errorlevel 1 (
    echo   SKIP: NSIS not found. Install NSIS to build the installer.
    echo   Download: https://nsis.sourceforge.io/Download
) else (
    makensis /V2 installer\OpenSheetET.nsi
    if errorlevel 1 (
        echo   WARNING: Installer build failed.
    ) else (
        echo   OK: Installer created.
    )
)

echo.
echo ╔══════════════════════════════════════════════════════════╗
echo ║  BUILD COMPLETE!                                         ║
echo ║                                                          ║
echo ║  Executable : %DEPLOY_DIR%\%APP_EXE%
echo ║  Run it     : %DEPLOY_DIR%\%APP_EXE%
echo ╚══════════════════════════════════════════════════════════╝
echo.
goto :eof

:error
echo.
echo ╔══════════════════════════════════════════════════════════╗
echo ║  BUILD FAILED — see errors above                         ║
echo ╚══════════════════════════════════════════════════════════╝
exit /b 1
