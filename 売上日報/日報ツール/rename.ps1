param([string]$Folder)

# ============================================================
#  Auto-rename scanned daily-report PDFs
#  Scan2026-10-17_204557.pdf  ->  (daily-report)10 (month)17 (day).pdf
#
#  This .ps1 is kept ASCII-ONLY on purpose. The Japanese target
#  name is built from Unicode code points so the file can never be
#  corrupted by a Shift-JIS / BOM misread when saved or copied.
# ============================================================

$ErrorActionPreference = 'Stop'

# --- 1) decide which folder holds the scans ---------------------------------
$candidates = @()
if (-not [string]::IsNullOrEmpty($Folder)) { $candidates += $Folder }
$candidates += (Join-Path $env:USERPROFILE 'Documents')
$candidates += (Join-Path $env:USERPROFILE 'OneDrive\Documents')

$target = $null
foreach ($c in $candidates) {
  if (Test-Path -LiteralPath $c) {
    $hit = @(Get-ChildItem -LiteralPath $c -Filter 'Scan*.pdf' -File -ErrorAction SilentlyContinue)
    if ($hit.Count -gt 0) { $target = $c; break }
  }
}
if ($null -eq $target) {
  foreach ($c in $candidates) { if (Test-Path -LiteralPath $c) { $target = $c; break } }
}
if ($null -eq $target) {
  Write-Host "ERROR: folder not found. Checked:"
  $candidates | ForEach-Object { Write-Host ("  " + $_) }
  exit 1
}
Write-Host ("Folder: " + $target)

# --- 2) Japanese name parts from code points (keeps source ASCII) -----------
$PREFIX = [char]0x58F2 + [char]0x4E0A + [char]0x65E5 + [char]0x5831   # uriage-nippou
$MONTH  = [string][char]0x6708                                        # month kanji
$DAY    = [string][char]0x65E5                                        # day kanji

# --- 3) rename every un-renamed Scan*.pdf by the date in its name -----------
$pat = '^Scan(\d{4})-(\d{2})-(\d{2})_'
$files = @(Get-ChildItem -LiteralPath $target -Filter 'Scan*.pdf' -File | Where-Object { $_.Name -match $pat })
if ($files.Count -eq 0) { Write-Host "No Scan*.pdf to rename."; exit 0 }

$done = 0
foreach ($f in $files) {
  if ($f.Name -match $pat) {
    $mo = [int]$Matches[2]
    $dy = [int]$Matches[3]
    $newName = $PREFIX + $mo + $MONTH + $dy + $DAY + '.pdf'
    $dest = Join-Path $target $newName
    if (Test-Path -LiteralPath $dest) {
      Write-Host ("SKIP (already exists): " + $newName + "   <- " + $f.Name)
    } else {
      Rename-Item -LiteralPath $f.FullName -NewName $newName
      Write-Host ("RENAMED: " + $f.Name + "  ->  " + $newName)
      $done++
    }
  }
}
Write-Host ("Done. renamed=" + $done)
