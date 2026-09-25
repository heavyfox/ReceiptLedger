param(
    [string]$Python = "python"
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$buildRoot = Join-Path $projectRoot "build"
$distRoot = Join-Path $projectRoot "dist"
$checksumFile = Join-Path $projectRoot "vendor\SHA256SUMS.txt"
$checksumLine = (Get-Content -LiteralPath $checksumFile -Raw).Trim() -split '\s+', 2
$wheelPath = Join-Path $projectRoot ("vendor\" + $checksumLine[1])
if ((Get-FileHash -LiteralPath $wheelPath -Algorithm SHA256).Hash.ToLowerInvariant() -ne $checksumLine[0]) {
    throw "HEIC wheel checksum mismatch"
}
& $Python -c "import pillow_heif; assert pillow_heif.__version__ == '1.8.0+receiptledger1', 'Install the decoder-only wheel from requirements.txt'; info = pillow_heif.libheif_info(); assert not info['HEIF'] and set(info['encoders']) <= {'mask'}, 'Unexpected HEIC encoder'"
if ($LASTEXITCODE -ne 0) { throw "HEIC dependency verification failed" }
& $Python -m PyInstaller --noconfirm --clean --distpath $distRoot --workpath $buildRoot `
    (Join-Path $projectRoot "ReceiptLedger.spec")
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed" }

# The runtime used in some development environments supplies Poppler's ICU DLLs.
# Qt on Windows expects the Windows ICU shim; avoid shipping unrelated copies.
$internal = Join-Path $distRoot "ReceiptLedger\_internal"
foreach ($name in @("icuuc.dll", "icudt78.dll")) {
    $target = Join-Path $internal $name
    if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target }
}
Write-Output (Join-Path $distRoot "ReceiptLedger\ReceiptLedger.exe")
