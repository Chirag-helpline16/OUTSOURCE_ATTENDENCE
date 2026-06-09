$ErrorActionPreference = "Stop"

$Root = (Resolve-Path (Split-Path -Parent $MyInvocation.MyCommand.Path)).Path
Set-Location $Root

function Remove-InProject {
    param([Parameter(Mandatory = $true)][string]$RelativePath)

    $target = Join-Path $Root $RelativePath
    if (-not (Test-Path $target)) {
        return
    }

    $resolved = (Resolve-Path $target).Path
    if (-not $resolved.StartsWith($Root, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to remove path outside project: $resolved"
    }

    Remove-Item -LiteralPath $resolved -Recurse -Force
}

$Venv = Join-Path $Root ".venv-build"
$Python = Join-Path $Venv "Scripts\python.exe"

if (-not (Test-Path $Python)) {
    if (Get-Command py -ErrorAction SilentlyContinue) {
        & py -3 -m venv $Venv
    } else {
        & python -m venv $Venv
    }
}

& $Python -m pip install --upgrade pip
& $Python -m pip install -r requirements.txt pyinstaller

Remove-InProject "build"
Remove-InProject "dist\OutsourceAttendance"

& $Python -m PyInstaller --noconfirm --clean OutsourceAttendance.spec

$Exe = Join-Path $Root "dist\OutsourceAttendance\OutsourceAttendance.exe"
if (-not (Test-Path $Exe)) {
    throw "Build finished, but the EXE was not found at $Exe"
}

Write-Host ""
Write-Host "Built successfully:"
Write-Host $Exe
Write-Host ""
Write-Host "Copy the whole dist\OutsourceAttendance folder to another PC."
Write-Host "Packaged EXE builds require MongoDB Atlas; SQLite fallback is disabled."
