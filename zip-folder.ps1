param(
  [Parameter(Mandatory=$true)][string]$InputDirectory,
  [Parameter(Mandatory=$true)][string]$OutputZip
)
$ErrorActionPreference='Stop'
$pdfs=Get-ChildItem -LiteralPath $InputDirectory -File -Filter '*.pdf'|Sort-Object Name
if(!$pdfs){throw 'No individual PDF letters were created.'}
Compress-Archive -LiteralPath $pdfs.FullName -DestinationPath $OutputZip -Force
