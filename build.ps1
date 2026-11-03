param(
    [string]$Python = "",
    [switch]$RefreshEngine,
    [switch]$Zip
)

$ErrorActionPreference = "Continue"
Set-Location $PSScriptRoot

if (-not $Python) {
    $candidates = @(
        "D:\Python312\python.exe",
        (Get-Command python -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty Source)
    )
    $Python = $candidates | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
}
if (-not $Python) {
    throw "Python not found. Pass -Python <path-to-python.exe>."
}
Write-Host "Using Python: $Python"

& $Python -c "import PySide6, pdf2docx, PyInstaller" 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "Missing build dependencies. Run: $Python -m pip install -r requirements.txt -r requirements-dev.txt"
}

if ($RefreshEngine -or -not (Test-Path "engine\node\node.exe")) {
    Write-Host "Assembling LibreOffice engine..."
    & $Python tools\prepare_engine.py
    if ($LASTEXITCODE -ne 0) { throw "Engine assembly failed." }
}

Write-Host "Running PyInstaller..."
& $Python -m PyInstaller --noconfirm --clean WenShuConverter.spec
if ($LASTEXITCODE -ne 0) { throw "PyInstaller build failed." }

$dist = Join-Path $PSScriptRoot "dist\WenShuConverter"
if (-not (Test-Path $dist)) { throw "Build output not found: $dist" }

Write-Host "Bundling portable conversion engine..."
if (Test-Path (Join-Path $dist "engine")) {
    Remove-Item (Join-Path $dist "engine") -Recurse -Force
}
Copy-Item "engine" (Join-Path $dist "engine") -Recurse -ErrorAction Stop
Copy-Item "README.md", "THIRD_PARTY_NOTICES.md", "使用说明.txt" $dist -Force -ErrorAction Stop
if (Test-Path "docs") {
    Copy-Item "docs" (Join-Path $dist "docs") -Recurse -Force -ErrorAction Stop
}

$size = (Get-ChildItem $dist -Recurse -File | Measure-Object Length -Sum).Sum / 1MB
Write-Host ("Build complete: {0} ({1:N0} MB)" -f $dist, $size)

if ($Zip) {
    $zipPath = Join-Path $PSScriptRoot "dist\WenShuConverter-Windows-x64.zip"
    if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
    Write-Host "Creating ZIP archive..."
    Compress-Archive -Path (Join-Path $dist "*") -DestinationPath $zipPath -CompressionLevel Optimal
    Write-Host "ZIP: $zipPath"
}
