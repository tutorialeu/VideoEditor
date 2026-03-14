@echo off
set VERSION=1.0.0.0
set DIST_DIR=dist
set BUILD_CMD="C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" /p:Configuration=Release

echo ========================================
echo   VideoEditor Release Packager v%VERSION%
echo ========================================

echo.
echo [1/4] Cleaning old artifacts...
if exist %DIST_DIR% rd /s /q %DIST_DIR%
mkdir %DIST_DIR%

echo [2/4] Building project in Release mode...
%BUILD_CMD%
if %ERRORLEVEL% neq 0 (
    echo.
    echo [ERROR] Build failed! Please check the errors above.
    pause
    exit /b %ERRORLEVEL%
)

echo [3/4] Gathering binaries and dependencies...
xcopy /y "bin\Release\*.exe" "%DIST_DIR%\"
xcopy /y "bin\Release\*.dll" "%DIST_DIR%\"
xcopy /y "bin\Release\*.config" "%DIST_DIR%\"

:: Ensure ffmpeg.exe is included if present
if exist "bin\Debug\ffmpeg.exe" (
    echo Copying FFmpeg...
    xcopy /y "bin\Debug\ffmpeg.exe" "%DIST_DIR%\"
)

echo [4/4] Finalizing...
:: Remove debugging files and unnecessary XMLs
del /q "%DIST_DIR%\*.pdb"
del /q "%DIST_DIR%\*.xml"

echo.
echo ========================================
echo   SUCCESS! Release ready in: %DIST_DIR%
echo ========================================
echo.
pause
