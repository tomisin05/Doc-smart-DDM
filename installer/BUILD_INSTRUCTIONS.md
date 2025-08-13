# Doc-smart Build Instructions

## Quick Start (PyInstaller - Recommended)

1. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Create icon (optional):**
   ```bash
   pip install Pillow
   python create_icon.py
   ```

3. **Build executable:**
   ```bash
   build_exe.bat
   ```
   
   Or manually:
   ```bash
   pyinstaller docsmart.spec
   ```

4. **Find your executable:**
   - Look in the `dist` folder for `Doc-smart.exe`

## Alternative: MSI Installer

1. **Build MSI installer:**
   ```bash
   build_installer.bat
   ```

2. **Find installer:**
   - Look in the `dist` folder for the `.msi` file

## Distribution Options

### Option 1: Single EXE File
- **Pros:** Easy to distribute, no installation needed
- **Cons:** Larger file size, slower startup
- **Use:** `pyinstaller --onefile --windowed docsmart.py`

### Option 2: Directory Distribution
- **Pros:** Faster startup, smaller main executable
- **Cons:** Multiple files to distribute
- **Use:** `pyinstaller --windowed docsmart.py`

### Option 3: MSI Installer
- **Pros:** Professional installation, Start Menu shortcuts, uninstaller
- **Cons:** More complex setup
- **Use:** `python setup.py bdist_msi`

## Advanced Options

### Code Signing (for distribution)
```bash
# After building, sign the executable
signtool sign /f certificate.pfx /p password /t http://timestamp.digicert.com Doc-smart.exe
```

### Auto-updater Integration
Consider adding auto-update functionality using libraries like:
- `pyupdater`
- `esky`

## Troubleshooting

### Common Issues:
1. **Missing modules:** Add to `hiddenimports` in spec file
2. **Large file size:** Use `--exclude-module` for unused packages
3. **Slow startup:** Use directory distribution instead of onefile

### Testing:
- Test on clean Windows machine without Python installed
- Check antivirus doesn't flag the executable
- Verify all features work in built version