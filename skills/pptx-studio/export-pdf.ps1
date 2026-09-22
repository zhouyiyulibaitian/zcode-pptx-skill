# export-pdf.ps1 -- pptx-studio fallback: export the deck PDF / page PNGs OUTSIDE the DSH sandbox.
#
# WHY THIS EXISTS
#   pptx-studio's `pdf` and `shots` subcommands print the deck HTML with headless
#   Chrome/Edge. Chrome's multi-process renderer needs mojo named pipes, and the DSH
#   workspace-write file sandbox denies them:
#       FATAL:mojo\public\cpp\platform\platform_channel.cc:108 Check failed: . : Access denied (0x5)
#   So those two subcommands always fail INSIDE a DSH session. That is sandbox policy,
#   not a skill bug: the very same HTML exported fine outside the sandbox (14 pages, 351 KB).
#
# HOW TO RUN (in a normal PowerShell window, NOT through a DSH tool call)
#   powershell -ExecutionPolicy Bypass -File export-pdf.ps1 -Deck "E:\work\deck.json"
#   powershell -ExecutionPolicy Bypass -File export-pdf.ps1 -Deck "E:\work\deck.json" -Shots
#   powershell -ExecutionPolicy Bypass -File export-pdf.ps1 -Html "E:\work\deck.html" -Out "E:\work\out.pdf"
#
# PARAMETERS
#   -Deck      deck.json; the script first runs studio.py render to produce the sibling .html
#   -Html      use an existing HTML instead of -Deck
#   -Out       output PDF path (default: the HTML path with a .pdf extension)
#   -Shots     also render every page to PNG (needs pypdfium2; drawn from pdf-toolkit's venv)
#   -Scale     PNG scale factor, default 2.0
#   -SkillDir  pptx-studio skill directory (default: this script's own directory)
#
# This file is deliberately ASCII-only so it parses identically under Windows PowerShell
# 5.1 and PowerShell 7+, regardless of BOM or console codepage.

param(
    [string]$Deck,
    [string]$Html,
    [string]$Out,
    [switch]$Shots,
    [double]$Scale = 2.0,
    [string]$SkillDir = $PSScriptRoot
)

$ErrorActionPreference = 'Stop'

if (-not $Deck -and -not $Html) { throw 'Need -Deck <deck.json> or -Html <file.html>' }
if ($Deck -and $Html) { throw 'Use either -Deck or -Html, not both' }

# --- interpreters and dependencies ----------------------------------------
# pptx-studio ships no venv and no requirements.txt; it uses the system Python's
# per-user site-packages (%APPDATA%\Python\Python314) for python-pptx / Pillow /
# pypdfium2 / lxml. Do NOT run it with `python -s` or PYTHONNOUSERSITE=1.
$GlobalPy = 'C:\Python314\python.exe'
$ShotsPy = 'C:\Users\zhouyi\.agents\skills\pdf-toolkit\.venv\Scripts\python.exe'

if (-not (Test-Path -LiteralPath $SkillDir)) { throw "SkillDir not found: $SkillDir" }
$SkillDir = (Resolve-Path -LiteralPath $SkillDir).Path
$StudioPy = Join-Path $SkillDir 'scripts\studio.py'

$ChromeCandidates = @(
    $env:CHROME_PATH,
    'C:\Program Files\Google\Chrome\Application\chrome.exe',
    'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
    'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
    'C:\Program Files\Microsoft\Edge\Application\msedge.exe'
) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }

if (-not $ChromeCandidates) { throw 'Chrome/Edge not found; set CHROME_PATH to the browser executable.' }
$Chrome = $ChromeCandidates[0]

if (-not (Test-Path -LiteralPath $GlobalPy)) { throw "System Python not found: $GlobalPy" }
if (-not (Test-Path -LiteralPath $StudioPy)) { throw "studio.py not found: $StudioPy" }

# --- 1) deck.json -> HTML --------------------------------------------------
if ($Deck) {
    $Deck = (Resolve-Path -LiteralPath $Deck).Path
    Write-Host "[1/3] render: $Deck"
    & $GlobalPy $StudioPy render --deck $Deck
    if ($LASTEXITCODE -ne 0) { throw "studio.py render failed (exit $LASTEXITCODE)" }
    $Html = [System.IO.Path]::ChangeExtension($Deck, '.html')
}

$Html = (Resolve-Path -LiteralPath $Html).Path
if (-not $Out) { $Out = [System.IO.Path]::ChangeExtension($Html, '.pdf') }
$OutFull = [System.IO.Path]::GetFullPath($Out)
$OutDir = Split-Path -Path $OutFull -Parent
if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Force -Path $OutDir | Out-Null }
if (Test-Path -LiteralPath $OutFull) { Remove-Item -LiteralPath $OutFull -Force }

# --- 2) headless browser -> PDF -------------------------------------------
Write-Host "[2/3] pdf: $Html -> $OutFull"
$url = 'file:///' + ($Html -replace '\\', '/')
$profileDir = Join-Path $env:TEMP ('dsh-pptx-pdf-' + [guid]::NewGuid().ToString('N').Substring(0, 8))

$headlessFlags = @('--headless=new', '--headless')
foreach ($flag in $headlessFlags) {
    # Start-Process -ArgumentList joins the array with spaces WITHOUT quoting, so any
    # value containing a space (both paths below, and the file:// URL when the deck lives
    # under a spaced path) must carry its own embedded quotes or Chrome sees extra
    # positional arguments and dies with "Multiple targets are not supported in headless mode".
    $chromeArgs = @(
        $flag
        '--disable-gpu'
        '--no-sandbox'
        ('"--user-data-dir=' + $profileDir + '"')
        '--no-pdf-header-footer'
        ('"--print-to-pdf=' + $OutFull + '"')
        ('"' + $url + '"')
    )
    Start-Process -FilePath $Chrome -ArgumentList $chromeArgs -Wait -NoNewWindow | Out-Null
    Start-Sleep -Milliseconds 500
    if ((Test-Path -LiteralPath $OutFull) -and (Get-Item -LiteralPath $OutFull).Length -gt 1200) { break }
    Write-Host "      $flag produced no file; trying the next headless mode"
}

if (-not (Test-Path -LiteralPath $OutFull)) {
    throw 'The browser produced no PDF. If you ran this inside a DSH tool call, the sandbox blocked Chrome named pipes -- run it in a normal PowerShell window instead.'
}

$kb = [math]::Round((Get-Item -LiteralPath $OutFull).Length / 1KB, 1)
Write-Host "      wrote $OutFull ($kb KB)"
Remove-Item -LiteralPath $profileDir -Recurse -Force -ErrorAction SilentlyContinue

# --- 3) optional: one PNG per page ----------------------------------------
if ($Shots) {
    if (-not (Test-Path -LiteralPath $ShotsPy)) { throw "PDF rasterizer interpreter not found: $ShotsPy" }
    $shotsDir = Join-Path (Split-Path -Path $OutFull -Parent) ([System.IO.Path]::GetFileNameWithoutExtension($OutFull) + '-shots')
    New-Item -ItemType Directory -Force -Path $shotsDir | Out-Null
    Write-Host "[3/3] shots: $shotsDir"
    # Rasterize with PyMuPDF: pdf-toolkit's venv has pymupdf but NOT pypdfium2,
    # and PyMuPDF's default 72 dpi matrix is exactly Scale=[$Scale].
    $pyCode = @'
import sys, pymupdf
pdf, outdir, scale = sys.argv[1], sys.argv[2], float(sys.argv[3])
doc = pymupdf.open(pdf)
for i, page in enumerate(doc, 1):
    page.get_pixmap(matrix=pymupdf.Matrix(scale, scale)).save(outdir + '/p%02d.png' % i)
print('      %d pages -> %s' % (doc.page_count, outdir))
'@
    & $ShotsPy -c $pyCode $OutFull $shotsDir $Scale
    if ($LASTEXITCODE -ne 0) { throw "page rendering failed (exit $LASTEXITCODE)" }
} else {
    Write-Host '[3/3] shots: skipped (pass -Shots to render page PNGs)'
}

Write-Host 'Done.'
