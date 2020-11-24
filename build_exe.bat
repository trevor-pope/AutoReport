@echo off && ^
venv\Scripts\activate && ^
pyinstaller build.spec ^
    --noconfirm ^
    --log-level=ERROR ^
    --icon=src\gui\assets\small_logo.ico && ^
echo on & ^
echo. & ^
echo. & ^
echo Finished building! You can run the executable found at "\dist\autoreport\autoreport.exe"
