echo Building Excel Diacritics Remover for Windows...
pyinstaller --onefile --windowed --name "ExcelDiacriticsRemover" --icon=icon.ico diacritics_remover.py
echo Build complete! Check the dist folder for ExcelDiacriticsRemover.exe
pause