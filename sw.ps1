# Thin wrapper so the SolidWorks bridge can be called from any shell:
#   .\sw.cmd sketch --plane front --shapes "rect 0 0 60 40"
# No param() block on purpose: it keeps "--out"/"--type" from being parsed as
# PowerShell parameters.
#
# Python discovery order:
#   1. $env:SW_PYTHON
#   2. .venv\Scripts\python.exe next to this script
#   3. py -3 / python3 / python, whichever can import pywin32
$env:PYTHONIOENCODING = 'utf-8'

$candidates = @()
if ($env:SW_PYTHON) { $candidates += , @($env:SW_PYTHON) }
$venv = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
if (Test-Path $venv) { $candidates += , @($venv) }
$candidates += , @('py', '-3')
$candidates += , @('python3')
$candidates += , @('python')

$python = $null
$prefix = @()
foreach ($c in $candidates) {
    $exe = $c[0]
    $extra = @()
    if ($c.Count -gt 1) { $extra = $c[1..($c.Count - 1)] }
    if (-not (Get-Command $exe -ErrorAction SilentlyContinue)) { continue }
    & $exe @extra -c "import win32com.client" 2>$null
    if ($LASTEXITCODE -eq 0) { $python = $exe; $prefix = $extra; break }
}

if (-not $python) {
    [Console]::Error.WriteLine(
        "No Python with pywin32 found. Install: py -3 -m pip install pywin32, " +
        "or set SW_PYTHON to a python.exe that has it.")
    exit 127
}

& $python @prefix (Join-Path $PSScriptRoot 'swbridge.py') @args
exit $LASTEXITCODE
