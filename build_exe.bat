@echo off
echo ================================
echo    Building TrackMyWorkout.exe
echo ================================
echo.

cd /d %~dp0

echo [1/3] Cleaning old build files...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [2/3] Building executable...
.venv\Scripts\pyinstaller.exe --noconfirm ^
    --onefile ^
    --windowed ^
    --name=TrackMyWorkout ^
    --clean ^
    --log-level=WARN ^
    src\app.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Build failed!
    pause
    exit /b 1
)

echo.
echo [3/3] Cleaning up...
if exist build rmdir /s /q build

echo.
echo ================================
echo    Build Complete!
echo ================================
echo.
echo EXE file location: %~dp0dist\TrackMyWorkout.exe
echo.
echo Note: All user data files (JSON) are excluded from the EXE.
echo       They will be created when the application runs.
echo.
pause
