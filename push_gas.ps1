$ErrorActionPreference = "Stop"
$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
$base = $PSScriptRoot

# find source txt (file name ends with GAS.txt, excluding backups)
$srcFile = Get-ChildItem -LiteralPath $base -Filter "*GAS.txt" | Where-Object { $_.Name -notmatch "backup" } | Select-Object -First 1
if (-not $srcFile) { Write-Host "ERROR: source *GAS.txt not found"; exit 1 }

# find the clasp project folder (the one that contains .clasp.json)
$dstDir = (Get-ChildItem -LiteralPath $base -Directory | Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName ".clasp.json") } | Select-Object -First 1).FullName
if (-not $dstDir) { Write-Host "ERROR: clasp project folder not found"; exit 1 }

# find the .js code file inside it
$dstFile = (Get-ChildItem -LiteralPath $dstDir -Filter "*.js" | Select-Object -First 1).FullName
if (-not $dstFile) { Write-Host "ERROR: code .js not found"; exit 1 }

Write-Host "[1/2] Copying latest code..."
Copy-Item -LiteralPath $srcFile.FullName -Destination $dstFile -Force

Write-Host "[2/2] Reflecting to Google Apps Script (please wait about 30 seconds)..."
Set-Location -LiteralPath $dstDir
$ErrorActionPreference = "Continue"
$out = cmd /c "npx --yes @google/clasp@2.4.2 push -f 2>&1"
$out = ($out | Out-String)

Write-Host ""
if ($out -match "Pushed") {
  Write-Host "==================================================="
  Write-Host "   SUCCESS !  Reflected to the spreadsheet GAS."
  Write-Host "   (You can close this window)"
  Write-Host "==================================================="
} else {
  Write-Host "###################################################"
  Write-Host "   FAILED.  Please show this whole screen to Claude."
  Write-Host "###################################################"
  Write-Host $out
}
