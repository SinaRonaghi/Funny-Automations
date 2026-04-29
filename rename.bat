@echo off
:: This line ensures the script runs in the folder where the .bat is saved
cd /d "%~dp0"

echo Current folder: %cd%
echo ---

:: Rename speech.wav to speech (0).wav
if exist "speech.wav" (
    ren "speech.wav" "speech (0).wav"
    echo [SUCCESS] Renamed speech.wav
) else (
    echo [NOT FOUND] speech.wav was not found in this folder.
)

:: Rename final.wav to zz.wav
if exist "final.wav" (
    ren "final.wav" "zz.wav"
    echo [SUCCESS] Renamed final.wav
) else (
    echo [NOT FOUND] final.wav was not found in this folder.
)

echo ---
echo Current files in folder:
dir /b *.wav

pause