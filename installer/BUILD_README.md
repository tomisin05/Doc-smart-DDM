# Doc-smart Build Instructions

## Quick Build
1. Run `build.bat` to create the executable
2. The executable will be created in `dist\Doc-smart.exe`

## Manual Build Steps
1. Install dependencies:
   ```
   pip install pyinstaller pywin32
   ```

2. Build executable:
   ```
   pyinstaller docsmart.spec
   ```

3. Create installer (optional):
   - Install Inno Setup from https://jrsoftware.org/isinfo.php
   - Run: `iscc installer.iss`
   - Installer will be created in `output\Doc-smart-Setup.exe`

## Files Created
- `dist\Doc-smart.exe` - Standalone executable
- `output\Doc-smart-Setup.exe` - Windows installer (if Inno Setup is used)

## Distribution
The executable in `dist\Doc-smart.exe` can be distributed as-is, or use the installer for a more professional installation experience.