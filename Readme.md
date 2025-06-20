# Excel Diacritics Remover - Setup and Build Instructions

## Overview
This application removes diacritics from Excel files, replacing accented characters with their non-accented equivalents. It provides a simple drag-and-drop interface for processing .xlsx, .xlsm, and .xls files.

## Features
- Cross-platform (Windows and macOS)
- Drag-and-drop interface
- Preserves Excel formulas (only modifies cell values)
- Creates a new file with "_fixed" suffix
- Supports all Excel file formats (.xlsx, .xlsm, .xls)

## Prerequisites
- Python 3.8 or higher
- pip (Python package manager)

## Setup Instructions

### 1. Install Python Dependencies
```bash
pip install -r requirements.txt
```

### 2. Test the Application
```bash
python diacritics_remover.py
```

## Building Executables

### For Windows (.exe)

1. Open Command Prompt or PowerShell
2. Navigate to the project directory
3. Run the build script:
```cmd
build_windows.bat
```

The .exe file will be created in the `dist` folder.

### For macOS (.dmg)

1. Open Terminal
2. Navigate to the project directory
3. Make the build script executable (if not already):
```bash
chmod +x build_macos.sh
```
4. Run the build script:
```bash
./build_macos.sh
```

The .dmg file will be created in the `dist` folder.

## Optional: Adding Icons

### Windows Icon
- Create or download a .ico file
- Name it `icon.ico` and place it in the project directory

### macOS Icon
- Create or download a .icns file
- Name it `icon.icns` and place it in the project directory

## Usage

1. Launch the application
2. Drag and drop an Excel file into the drop zone
3. Click the "Fix" button
4. The processed file will be saved in the same directory as the original with "_fixed" added to the filename

## Troubleshooting

### tkinterdnd2 Issues on macOS
If you encounter issues with tkinterdnd2 on macOS, you may need to install it from source:
```bash
pip install tkinterdnd2-universal
```

### PyInstaller Issues
If PyInstaller fails, try:
1. Clear the build cache: `pyinstaller --clean`
2. Update PyInstaller: `pip install --upgrade pyinstaller`
3. Use virtual environment to avoid conflicts

## Diacritics Mapping
The application uses the same character mapping as your VBA code, including:
- À, Á, Â, Ã, Ä, Å → A
- È, É, Ê, Ë → E
- Ì, Í, Î, Ï → I
- Ò, Ó, Ô, Õ, Ö, Ø → O
- Ù, Ú, Û, Ü → U
- And many more...

## Notes
- The application creates a new file rather than modifying the original
- Formulas in cells are preserved (not modified)
- All worksheets in the Excel file are processed