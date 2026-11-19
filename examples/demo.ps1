# Builds the sample plate: 60x40x10 block, D16 through hole, 4x R3 vertical fillets.
# Run from the repo root:  powershell -ExecutionPolicy Bypass -File examples\demo.ps1
$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
$sw   = Join-Path $root "sw.cmd"
$out  = Join-Path $root "out"

function Step {
    param([string]$Label, [string[]]$Argv)
    Write-Host "== $Label" -ForegroundColor Cyan
    & $sw @Argv | Out-Host
    if ($LASTEXITCODE -ne 0) { throw "$Label failed (exit $LASTEXITCODE)" }
}

Step "status"  @("status")
Step "new"     @("new", "--type", "part")
Step "plate"   @("sketch",  "--plane", "front", "--shapes", "rect 0 0 60 40")
Step "extrude" @("extrude", "--depth", "10")
Step "hole"    @("sketch",  "--plane", "front", "--shapes", "circle 0 0 8")
Step "cut"     @("cut", "--dir")
Step "fillet"  @("fillet", "--radius", "3", "--edges", "vertical")
Step "info"    @("info")
Step "save"    @("save",   "--path", (Join-Path $out "plate.SLDPRT"))
Step "export"  @("export", "--out",  (Join-Path $out "plate.step"))
Step "close"   @("close")

Write-Host "expected: volume 21912.1 mm3, bbox 60 x 40 x 10" -ForegroundColor Green
