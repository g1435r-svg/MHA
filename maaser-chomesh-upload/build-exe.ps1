$node = Get-Command node -ErrorAction SilentlyContinue
$npm = Get-Command npm -ErrorAction SilentlyContinue

if (-not $node -or -not $npm) {
  Write-Host "Node.js/npm not found. Install Node.js 20+ and run again." -ForegroundColor Red
  exit 1
}

Write-Host "Installing dependencies..." -ForegroundColor Yellow
npm install
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Write-Host "Building portable EXE..." -ForegroundColor Green
npm run build:exe
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Write-Host "Build completed. See dist folder." -ForegroundColor Green
