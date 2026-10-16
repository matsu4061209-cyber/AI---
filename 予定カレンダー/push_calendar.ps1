$ErrorActionPreference = "Stop"
$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
$base = $PSScriptRoot

# destination = subfolder that has .clasp.json (ignore folders starting with "_")
$dstDir = (Get-ChildItem -LiteralPath $base -Directory | Where-Object { ($_.Name -notlike "_*") -and (Test-Path -LiteralPath (Join-Path $_.FullName ".clasp.json")) } | Select-Object -First 1).FullName
if (-not $dstDir) { Write-Host "ERROR: clasp folder not found"; exit 1 }

# source = subfolder that has a .gs code file (the editable originals)
$srcDir = (Get-ChildItem -LiteralPath $base -Directory | Where-Object { ($_.Name -notlike "_*") -and ($_.FullName -ne $dstDir) -and (Get-ChildItem -LiteralPath $_.FullName -Filter "*.gs" -ErrorAction SilentlyContinue) } | Select-Object -First 1).FullName
if (-not $srcDir) { Write-Host "ERROR: source folder not found"; exit 1 }

# copy latest code into the clasp folder
$srcGs = (Get-ChildItem -LiteralPath $srcDir -Filter "*.gs" | Select-Object -First 1).FullName
$dstJs = (Get-ChildItem -LiteralPath $dstDir -Filter "*.js" | Select-Object -First 1).FullName
Copy-Item -LiteralPath $srcGs -Destination $dstJs -Force
Copy-Item -LiteralPath (Join-Path $srcDir "index.html") -Destination (Join-Path $dstDir "index.html") -Force

Write-Host "Pushing code to Google Apps Script..."
Set-Location -LiteralPath $dstDir
$ErrorActionPreference = "Continue"
$out = cmd /c "npx --yes @google/clasp@2.4.2 push -f 2>&1"
$out = ($out | Out-String)
Write-Host $out

if ($out -match "Pushed") {
  Write-Host "=================================================="
  Write-Host "  CODE UPDATED on the server."
  Write-Host "  Final step (in the WORK Google account):"
  Write-Host "   Apps Script > Deploy > Manage deployments"
  Write-Host "   > pencil (edit) > Version = New version > Deploy"
  Write-Host "  The URL stays the same."
  Write-Host "=================================================="
} else {
  Write-Host "##################################################"
  Write-Host "  FAILED. Please show this whole screen to Claude."
  Write-Host "##################################################"
}
