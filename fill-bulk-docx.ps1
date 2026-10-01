param(
  [Parameter(Mandatory=$true)][string]$InputPath,
  [Parameter(Mandatory=$true)][string]$ManifestPath
)
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $MyInvocation.MyCommand.Path
$parsed=Get-Content -Raw -LiteralPath $InputPath|ConvertFrom-Json
$items=@()
foreach($entry in $parsed){$items+=$entry}
$work=Split-Path -Parent ([IO.Path]::GetFullPath($ManifestPath))
$records=New-Object System.Collections.Generic.List[object]
$errors=New-Object System.Collections.Generic.List[object]
$sequence=0
foreach($item in $items){
  $sequence++
  $json=[IO.Path]::GetFullPath((Join-Path $work ("bulk-payload-{0:000}.json"-f$sequence)))
  try{
    $docxPath=[IO.Path]::GetFullPath([string]$item.docxPath)
    $item.payload|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $json -Encoding UTF8
    & (Join-Path $root 'generate-letter.ps1') -DataPath $json -OutputPath $docxPath
    $records.Add([ordered]@{
      name=[string]$item.payload.name
      employeeId=[string]$item.payload.employeeId
      docxPath=$docxPath
      pdfPath=[string]$item.pdfPath
      fileName=[string]$item.fileName
      salaryAnnexurePath=[string]$item.salaryAnnexurePath
    })
  }catch{
    $errors.Add([ordered]@{
      name=[string]$item.payload.name
      employeeId=[string]$item.payload.employeeId
      error=$_.Exception.Message
      dataPath=$json
      outputPath=[string]$item.docxPath
    })
  }
}
if(!$records.Count){
  $detail=if($errors.Count){$errors[0]|ConvertTo-Json -Compress}else{'No employee rows were supplied.'}
  throw "No letters were generated. $detail"
}
[ordered]@{
  generated=$records.Count
  failed=$errors.Count
  records=$records
  errors=$errors
}|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $ManifestPath -Encoding UTF8
