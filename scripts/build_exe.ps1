$ErrorActionPreference = "Stop"

Write-Host "Installing dependencies..."
python -m pip install -r requirements.txt

Write-Host "Cleaning old build outputs..."
if (Test-Path ".\build") {
    Remove-Item ".\build" -Recurse -Force
}
if (Test-Path ".\dist") {
    Remove-Item ".\dist" -Recurse -Force
}

Write-Host "Building executable with PyInstaller..."
pyinstaller --noconfirm --clean ".\packaging\build_pizza.spec"

$exePath = ".\dist\PizzaTool\PizzaTool.exe"
if (Test-Path $exePath) {
    Write-Host "Build succeeded: $exePath"
} else {
    throw "Build failed: executable not found at $exePath"
}
