@echo off
echo Building SnapCherry...
gcc SnapCherry.c -o SnapCherry.exe -mwindows -luser32 -lgdi32 -lole32 -lwindowscodecs -lmsimg32 -lshlwapi -lshell32 -loleaut32 -luuid
if %ERRORLEVEL% EQU 0 (
    echo Build successful: SnapCherry.exe
) else (
    echo Build failed.
)
pause
