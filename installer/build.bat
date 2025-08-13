@echo off
echo Building Doc-smart executable...

REM Install required packages
pip install pyinstaller pywin32

REM Build executable using spec file
pyinstaller docsmart.spec

REM Check if build was successful
if exist "dist\Doc-smart.exe" (
    echo.
    echo Build successful! Executable created at: dist\Doc-smart.exe
    echo.
    echo To create installer, run: iscc installer.iss
    echo (Requires Inno Setup to be installed)
) else (
    echo.
    echo Build failed! Check for errors above.
)

pause