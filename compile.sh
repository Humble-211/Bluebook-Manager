#!/bin/bash

echo "Cleaning previous builds..."
rm -rf build dist

echo "Compiling Bluebook Manager..."
"C:/Users/hmai/AppData/Local/Packages/PythonSoftwareFoundation.Python.3.12_qbz5n2kfra8p0/LocalCache/local-packages/Python312/Scripts/pyinstaller.exe" BluebookManager.spec --clean -y

echo ""
echo "Compilation finished! The executable is located in the 'dist' folder."
