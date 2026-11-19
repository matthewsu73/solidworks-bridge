# Surface workflow demo: build a sheet, thicken it into a solid, then a lofted sheet.
# Run from the repo root:  powershell -ExecutionPolicy Bypass -File examples\surface.ps1
$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
$sw   = Join-Path $root "sw.cmd"

function Step {
    param([string]$Label, [string[]]$Argv)
    Write-Host "== $Label" -ForegroundColor Cyan
    & $sw @Argv | Out-Host
    if ($LASTEXITCODE -ne 0) { throw "$Label failed (exit $LASTEXITCODE)" }
}

# 1. sheet 60x40 extruded 25 -> area 5000 mm2, then thicken 3 -> solid 15900 mm3
Step "new"       @("new", "--type", "part")
Step "profile"   @("sketch", "--plane", "front", "--shapes", "rect 0 0 60 40")
Step "sheet"     @("surface-extrude", "--depth", "25")
Step "info"      @("info")
Step "thicken"   @("thicken", "--thickness", "3", "--faces", "all")
Step "info2"     @("info")
Step "close1"    @("close")

# 2. loft: rectangle on the front plane -> circle on an offset plane
Step "new2"      @("new", "--type", "part")
Step "sk1"       @("sketch", "--plane", "front", "--shapes", "rect 0 0 60 40")
Step "plane"     @("plane-offset", "--base", "front", "--distance", "40")
Step "sk2"       @("sketch", "--plane", "基准面1", "--shapes", "circle 0 0 15")
Step "loft"      @("surface-loft", "--sketches", "草图1,草图2")
Step "info3"     @("info")

# 3. sweep: circle on the front plane swept along a line on the right plane
Step "new3"      @("new", "--type", "part")
Step "sk3"       @("sketch", "--plane", "front", "--shapes", "circle 0 0 5")
Step "sk4"       @("sketch", "--plane", "right", "--shapes", "line 0 0 40 0")
Step "sweep"     @("surface-sweep", "--outline", "草图1", "--path", "草图2")
Step "info4"     @("info")

# 4. blank 120x120x25 minus a D40x10 boss -> 347433.63 mm3 (mould-style subtract)
Step "new4"      @("new", "--type", "part")
Step "sk5"       @("sketch", "--plane", "front", "--shapes", "rect 0 0 120 120")
Step "blank"     @("extrude", "--depth", "25")
Step "sk6"       @("sketch", "--plane", "front", "--shapes", "circle 0 0 20")
Step "boss"      @("extrude", "--depth", "10", "--no-merge")
Step "combine"   @("combine", "--operation", "subtract", "--target", "largest")
Step "info5"     @("info")

# 5. mould cavity from a thickened loft shell: 120x120x25 blank, loft 80x80(z=3) -> D50(z=20)
#    thicken 3 --no-merge keeps the shell a separate body so combine can subtract it
#    -> 341300.26 mm3, back to 1 solid body
Step "new5"      @("new", "--type", "part")
Step "sk7"       @("sketch", "--plane", "front", "--shapes", "rect 0 0 120 120")
Step "blank2"    @("extrude", "--depth", "25")
Step "plane2"    @("plane-offset", "--base", "front", "--distance", "3")
Step "sk8"       @("sketch", "--plane", "基准面1", "--shapes", "rect 0 0 80 80")
Step "plane3"    @("plane-offset", "--base", "基准面1", "--distance", "17")
Step "sk9"       @("sketch", "--plane", "基准面2", "--shapes", "circle 0 0 25")
Step "loft2"     @("surface-loft", "--sketches", "草图2,草图3")
Step "shell"     @("thicken", "--thickness", "3", "--faces", "all", "--no-merge")
Step "subtract"  @("combine", "--operation", "subtract", "--target", "largest")
Step "info6"     @("info")

Write-Host "expected: thicken 15900 mm3 / loft bbox 60x40x40 / sweep bbox 10x10x40 / subtract 347433.63 mm3 / cavity 341300.26 mm3" -ForegroundColor Green
