param([Parameter(Mandatory=$true)][string]$MailPath)
$ErrorActionPreference = "Stop"
$mailData = Get-Content -Raw -LiteralPath $MailPath | ConvertFrom-Json
$outlook = $null
$message = $null
try {
  $outlook = New-Object -ComObject Outlook.Application
  $message = $outlook.CreateItem(0)
  $message.To = [string]$mailData.to
  $message.Subject = [string]$mailData.subject
  $message.HTMLBody = [string]$mailData.html
  if ($mailData.attachmentPath -and (Test-Path -LiteralPath ([string]$mailData.attachmentPath))) {
    [void]$message.Attachments.Add([string]$mailData.attachmentPath)
  }
  $message.Send()
} finally {
  if ($message) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($message) }
  if ($outlook) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
