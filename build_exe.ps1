param(
    [string]$Python = "python"
)

$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$EntryPoint = Join-Path $ProjectDir "landt_workbench.py"

if (-not (Test-Path -LiteralPath $EntryPoint)) {
    throw "Cannot find entry point: $EntryPoint"
}

Push-Location $ProjectDir
try {
    & $Python -m PyInstaller `
        --noconfirm `
        --clean `
        --onefile `
        --windowed `
        --name "LAND-Data-Workbench" `
        --hidden-import "PIL._tkinter_finder" `
        --hidden-import "matplotlib.backends.backend_agg" `
        --hidden-import "matplotlib.backends.backend_pdf" `
        --hidden-import "matplotlib.backends.backend_ps" `
        --hidden-import "matplotlib.backends.backend_svg" `
        "landt_workbench.py"

    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller failed with exit code $LASTEXITCODE"
    }

    $ExePath = Join-Path $ProjectDir "dist\LAND-Data-Workbench.exe"
    if (-not (Test-Path -LiteralPath $ExePath)) {
        throw "Build completed without the expected executable: $ExePath"
    }
    Write-Host "Built: $ExePath"
}
finally {
    Pop-Location
}
