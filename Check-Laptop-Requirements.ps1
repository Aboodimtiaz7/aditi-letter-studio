$ErrorActionPreference = "SilentlyContinue"

function Show-Result([string]$Name, [bool]$Available, [string]$Details) {
  $status = if ($Available) { "READY" } else { "MISSING" }
  $colour = if ($Available) { "Green" } else { "Yellow" }
  Write-Host ("{0,-22} {1,-8} {2}" -f $Name, $status, $Details) -ForegroundColor $colour
}

Write-Host ""
Write-Host "ADITI Letter Studio - Laptop Requirements" -ForegroundColor Cyan
Write-Host "------------------------------------------------------------"

$node = Get-Command node.exe
$nodeVersion = if ($node) { (& $node.Source --version) } else { "Install the organisation-approved Node.js LTS package." }
$wordAvailable = $null -ne [type]::GetTypeFromProgID("Word.Application")
$excelAvailable = $null -ne [type]::GetTypeFromProgID("Excel.Application")
$outlookAvailable = $null -ne [type]::GetTypeFromProgID("Outlook.Application")
$browserPath = @(
  "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
  "C:\Program Files\Microsoft\Edge\Application\msedge.exe",
  "C:\Program Files\Google\Chrome\Application\chrome.exe",
  "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
$browserAvailable = $null -ne $browserPath

Show-Result "Windows PowerShell" ($PSVersionTable.PSVersion.Major -ge 5) ("Version " + $PSVersionTable.PSVersion)
Show-Result "Node.js" ($null -ne $node) $nodeVersion
Show-Result "Edge or Chrome" $browserAvailable $(if ($browserAvailable) { Split-Path -Leaf $browserPath } else { "Install Microsoft Edge or Google Chrome" })
Show-Result "Microsoft Excel" $excelAvailable $(if ($excelAvailable) { "Optional desktop app detected" } else { "Optional; only for a separately exported salary annexure" })
Show-Result "Microsoft Word" $wordAvailable $(if ($wordAvailable) { "Optional desktop app detected" } else { "Optional; used only to save ODT as DOCX" })
Show-Result "Microsoft Outlook" $outlookAvailable $(if ($outlookAvailable) { "Optional email app detected" } else { "Optional; required only for Outlook email sending" })

Write-Host "------------------------------------------------------------"
if ($node -and $browserAvailable) {
  Write-Host "This laptop is ready for local letter generation." -ForegroundColor Green
  exit 0
}

Write-Host "Install the missing required items, restart Windows, and run this check again." -ForegroundColor Yellow
exit 1
