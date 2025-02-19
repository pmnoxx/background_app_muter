# Stop on any error
$ErrorActionPreference = "Stop"

# Define paths
$destinationPath = "D:\repos\game_mods\apps\App Muter"
$distPath = ".\dist\App Muter"

# Create destination directory if it doesn't exist
if (-not (Test-Path $destinationPath)) {
    New-Item -ItemType Directory -Path $destinationPath
}

# Clean previous build artifacts but keep dist
Write-Host "Cleaning previous build..."
if (Test-Path ".\build") { Remove-Item ".\build" -Recurse -Force }

# Build with PyInstaller
Write-Host "Building with PyInstaller..."
pyinstaller app_muter.spec --noconfirm

# Check if build was successful
if (-not (Test-Path $distPath)) {
    Write-Host "Build failed! Dist folder not found." -ForegroundColor Red
    exit 1
}

# Copy to destination
Write-Host "Copying to $destinationPath..."
Copy-Item -Path "$distPath\*" -Destination $destinationPath -Recurse -Force

Write-Host "Build completed successfully!" -ForegroundColor Green
Write-Host "Files copied to: $destinationPath" -ForegroundColor Green 