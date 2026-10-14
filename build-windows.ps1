$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
if ($env:OS -ne 'Windows_NT') { throw 'Build this package on Windows.' }
function Run-Checked {
    param([string]$Exe, [string[]]$Arguments)
    & $Exe @Arguments
    if ($LASTEXITCODE -ne 0) { throw "Command failed: $Exe $Arguments" }
}
if (-not (Test-Path '.venv\Scripts\python.exe')) {
    if (Get-Command py -ErrorAction SilentlyContinue) {
        Run-Checked 'py' @('-3.12', '-m', 'venv', '.venv')
    } elseif (Get-Command python -ErrorAction SilentlyContinue) {
        Run-Checked 'python' @('-m', 'venv', '.venv')
    } else { throw 'Install Python 3.12 (64-bit) from python.org first.' }
}
$python = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
Run-Checked $python @('-m', 'pip', 'install', '-r', 'requirements-build.txt')
Run-Checked $python @('-m', 'pip', 'install', '--no-build-isolation', '--no-deps', '-e', '.')
$env:QT_QPA_PLATFORM = 'offscreen'
Run-Checked $python @('-m', 'pytest', '-q')
Remove-Item Env:\QT_QPA_PLATFORM
Run-Checked $python @('-m', 'PyInstaller', '--noconfirm', '--clean', 'QingBiao.spec')
Write-Host 'Done: dist\QingBiao.exe (single file, no installation required)' -ForegroundColor Green
