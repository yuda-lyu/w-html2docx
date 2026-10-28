@echo off
rem Build htmlToDocx.exe (one-file) from htmlToDocx.spec and output it to ..\bin\ (external executables live in bin/, not in src/).
rem Requires: pip install pyinstaller pywin32
pushd "%~dp0"
python -m PyInstaller htmlToDocx.spec --noconfirm --clean --distpath ..\bin
popd
pause
