cat > build_macos.sh << 'EOF'
#!/bin/bash

echo "Building Excel Diacritics Remover for macOS..."

# Build the .app bundle
pyinstaller --onefile --windowed \
    --name "ExcelDiacriticsRemover" \
    --icon=icon.icns \
    --osx-bundle-identifier "com.yourcompany.exceldiacriticsremover" \
    diacritics_remover.py

# Create DMG
echo "Creating DMG..."
cd dist

# Create a temporary directory for DMG contents
mkdir -p dmg_temp
cp -R ExcelDiacriticsRemover.app dmg_temp/

# Create the DMG
hdiutil create -volname "Excel Diacritics Remover" \
    -srcfolder dmg_temp \
    -ov -format UDZO \
    ExcelDiacriticsRemover.dmg

# Clean up
rm -rf dmg_temp

echo "Build complete! Check the dist folder for ExcelDiacriticsRemover.dmg"
EOF

chmod +x build_macos.sh