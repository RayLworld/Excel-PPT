$ErrorActionPreference = "Stop"

Write-Host "Step 1/2: Building executable with PyInstaller..."
powershell -ExecutionPolicy Bypass -File ".\scripts\build_exe.ps1"

$issPath = ".\packaging\PizzaTool.iss"
if (-not (Test-Path $issPath)) {
    throw "Inno Setup script not found: $issPath"
}

$isccCandidates = @(
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
    "${env:ProgramFiles}\Inno Setup 6\ISCC.exe",
    "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
)

$isccPath = $null
foreach ($candidate in $isccCandidates) {
    if (Test-Path $candidate) {
        $isccPath = $candidate
        break
    }
}

if (-not $isccPath) {
    throw "ISCC.exe not found. Please install Inno Setup 6 first: https://jrsoftware.org/isinfo.php"
}

Write-Host "Step 2/2: Building installer with Inno Setup..."
& $isccPath $issPath

$installerPath = ".\dist\installer\PizzaTool-Setup.exe"
if (Test-Path $installerPath) {
    Write-Host "Installer build succeeded: $installerPath"
} else {
    throw "Installer build failed: $installerPath not found"
}
