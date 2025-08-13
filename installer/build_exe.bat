@echo off
echo Building Doc-smart executable...
pip install -r requirements.txt
pyinstaller --onefile --windowed --name "Doc-smart" --icon=icon.ico docsmart.py
echo Build complete! Check the 'dist' folder for Doc-smart.exe
pause