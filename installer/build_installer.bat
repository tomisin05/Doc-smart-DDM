@echo off
echo Installing cx_Freeze...
pip install cx_Freeze

echo Building executable...
python setup.py build

echo Creating MSI installer...
python setup.py bdist_msi

echo Build complete! Check the 'dist' folder for the installer.
pause