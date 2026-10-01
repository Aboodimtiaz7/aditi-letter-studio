param([string]$OpenPath = "/")
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here
$nodeCommand = Get-Command node.exe -ErrorAction SilentlyContinue
if (-not $nodeCommand) {
  $installedNode = "C:\Program Files\nodejs\node.exe"
  if (Test-Path -LiteralPath $installedNode) {
    $nodeCommand = Get-Item -LiteralPath $installedNode
  }
}
if (-not $nodeCommand) {
  Write-Host ""
  Write-Host "Node.js is required before ADITI Letter Studio can start." -ForegroundColor Yellow
  Write-Host "Please install the organisation-approved Node.js LTS version, restart Windows, and run this shortcut again."
  Read-Host "Press Enter to close"
  exit 1
}
$nodeExecutable = if ($nodeCommand.Source) { $nodeCommand.Source } else { $nodeCommand.FullName }
if (-not (Test-Path "node_modules")) {
  $npmCommand = Get-Command npm.cmd -ErrorAction SilentlyContinue
  if (-not $npmCommand) {
    Write-Host "npm is missing. Reinstall the organisation-approved Node.js LTS package." -ForegroundColor Yellow
    Read-Host "Press Enter to close"
    exit 1
  }
  & $npmCommand.Source install
  if ($LASTEXITCODE -ne 0) { Read-Host "Installation failed. Check the internet connection or contact IT. Press Enter to close"; exit 1 }
}
if (-not (Test-Path "dist\index.html")) { npm.cmd run build; if ($LASTEXITCODE -ne 0) { Read-Host "Build failed. Press Enter to close"; exit 1 } }
$listenerPids = @()
$existing = Get-NetTCPConnection -LocalPort 5180 -State Listen -ErrorAction SilentlyContinue
if ($existing) { $listenerPids += @($existing | ForEach-Object { $_.OwningProcess }) }
if (-not $listenerPids) {
  $netstatMatches = netstat -ano -p tcp | Select-String -Pattern ':5180\s+.*LISTENING\s+(\d+)\s*$'
  foreach ($match in $netstatMatches) { if ($match.Matches.Count) { $listenerPids += [int]$match.Matches[0].Groups[1].Value } }
}
foreach ($listenerPid in @($listenerPids | Select-Object -Unique)) {
  $listener = Get-Process -Id $listenerPid -ErrorAction SilentlyContinue
  if ($listener -and $listener.ProcessName -eq "node") { Stop-Process -Id $listener.Id -Force; Start-Sleep -Milliseconds 500 }
  elseif ($listener) { Read-Host "Port 5180 is being used by another application. Please close it and press Enter"; exit 1 }
}
Start-Process -FilePath $nodeExecutable -ArgumentList "server.js" -WorkingDirectory $here -WindowStyle Hidden
$ready=$false
for($i=0;$i -lt 60;$i++){
  try { $health=Invoke-RestMethod "http://127.0.0.1:5180/api/health" -TimeoutSec 1; if($health.ok){$ready=$true;break} } catch {}
  Start-Sleep -Milliseconds 500
}
if(-not $OpenPath.StartsWith("/")){$OpenPath="/$OpenPath"}
if($ready){Start-Process "http://127.0.0.1:5180$OpenPath"}else{Read-Host "Letter Studio could not start. Press Enter to close";exit 1}
