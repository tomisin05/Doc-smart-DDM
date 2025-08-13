from cx_Freeze import setup, Executable
import sys

# Dependencies are automatically detected, but it might need fine tuning.
build_options = {
    'packages': ['tkinter', 'win32com.client'],
    'excludes': ['test', 'unittest'],
    'include_files': []
}

base = 'Win32GUI' if sys.platform == 'win32' else None

executables = [
    Executable('docsmart.py', 
              base=base, 
              target_name='Doc-smart.exe',
              icon='icon.ico')
]

setup(
    name='Doc-smart',
    version='1.0.0',
    description='Debate Document Manager',
    author='Your Name',
    options={'build_exe': build_options},
    executables=executables
)