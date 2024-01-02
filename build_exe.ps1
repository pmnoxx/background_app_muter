# Set parameters
$pythonApp = "app_muter.py"
$exeName = "app_muter.exe" 

# Create dist folder if it doesn't exist
New-Item -ItemType Directory -Path .\dist

# Install pyinstaller if not already installed
pip install pyinstaller 

# Convert Python app to EXE in dist folder 
pyinstaller --onefile --distpath .\dist $pythonApp

# Move the EXE to the current folder
# Move-Item -Path .\dist\$exeName -Destination .\ 

# Remove old build folder
# Remove-Item .\dist -Recurse