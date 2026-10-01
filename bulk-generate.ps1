param(
  [Parameter(Mandatory=$true)][string]$BaseDataPath,
  [Parameter(Mandatory=$true)][string]$EmployeePath,
  [Parameter(Mandatory=$true)][string]$OutputDirectory,
  [Parameter(Mandatory=$true)][string]$ManifestPath
)
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $root 'word-automation.ps1')
$base=Get-Content -Raw -LiteralPath $BaseDataPath | ConvertFrom-Json
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Path $OutputDirectory -Force|Out-Null

function Key($value){([string]$value).ToLowerInvariant()-replace'[^a-z0-9]',''}
function Pick($row,[string[]]$aliases){
  foreach($alias in $aliases){
    $key=Key $alias
    if($row.ContainsKey($key)-and $null-ne$row[$key]-and -not [string]::IsNullOrWhiteSpace([string]$row[$key])){return $row[$key]}
  }
  return $null
}
function DateText($value){
  if($null-eq$value-or[string]::IsNullOrWhiteSpace([string]$value)){return ''}
  if($value-is[double]-or$value-is[int]){return [datetime]::FromOADate([double]$value).ToString('yyyy-MM-dd')}
  try{return (Get-Date $value).ToString('yyyy-MM-dd')}catch{return [string]$value}
}
function BoolValue($value){
  if($value-is[bool]){return $value}
  return ([string]$value).Trim().ToLowerInvariant()-in@('yes','y','true','1')
}
function NumberValue($value){
  if($null-eq$value-or[string]::IsNullOrWhiteSpace([string]$value)){return $null}
  $number=0.0
  if([double]::TryParse(([string]$value-replace',',''),[ref]$number)){return $number}
  return $null
}
function Safe($value){
  $safe=([string]$value)-replace'[\\/:*?"<>|]','-'
  $safe=$safe.Trim().TrimEnd('.')
  if(!$safe){return 'Employee'}
  return $safe
}
function Put($target,$row,[string]$property,[string[]]$aliases,[string]$kind='text'){
  $value=Pick $row $aliases
  if($null-eq$value){return}
  if($kind-eq'date'){$target[$property]=DateText $value}
  elseif($kind-eq'number'){$number=NumberValue $value;if($null-ne$number){$target[$property]=$number}}
  elseif($kind-eq'bool'){$target[$property]=BoolValue $value}
  else{$target[$property]=[string]$value}
}
function New-Payload($baseObject){
  $payload=[ordered]@{}
  foreach($property in $baseObject.PSObject.Properties){
    if($property.Name-notin@('employeeExcelBase64','employeeExcelName','customTemplateBase64')){$payload[$property.Name]=$property.Value}
  }
  return $payload
}
function Add-Salary($payload,$sheet,$excel,[bool]$isContractor,[bool]$exportAnnexure,[string]$annexurePath){
  $gross=NumberValue $payload.grossCtc
  if($null-eq$gross-or$gross-le0){throw 'Gross CTC is required for this salary letter.'}
  $sheet.Range('C4').Value2=[string]$payload.employeeId
  $sheet.Range('C5').Value2=[string]$payload.name
  $sheet.Range('C6').Value2=[string]$payload.designation
  $sheet.Range('C7').Value2=if(BoolValue $payload.pfAtActualBasic){'Yes'}else{'No'}
  $sheet.Range('C17').Value2=[double]$gross
  $sheet.Calculate()
  $payload.grossCtc=[math]::Round([double]$sheet.Range('C17').Value2)
  $payload.basic=[math]::Round([double]$sheet.Range('C10').Value2)
  $payload.basicMonthly=[math]::Round([double]$sheet.Range('D10').Value2)
  $payload.hra=[math]::Round([double]$sheet.Range('C11').Value2)
  $payload.hraMonthly=[math]::Round([double]$sheet.Range('D11').Value2)
  $payload.bonus=[math]::Round([double]$sheet.Range('C12').Value2)
  $payload.bonusMonthly=[math]::Round([double]$sheet.Range('D12').Value2)
  $payload.special=[math]::Round([double]$sheet.Range('C13').Value2)
  $payload.specialMonthly=[math]::Round([double]$sheet.Range('D13').Value2)
  $payload.grossSalary=[math]::Round([double]$sheet.Range('C15').Value2)
  $payload.grossSalaryMonthly=[math]::Round([double]$sheet.Range('D15').Value2)
  $payload.pf=[math]::Round([double]$sheet.Range('C16').Value2)
  $payload.pfMonthly=[math]::Round([double]$sheet.Range('D16').Value2)
  $payload.medical=[math]::Round([double]$sheet.Range('C20').Value2)
  $payload.medicalMonthly=[math]::Round([double]$sheet.Range('D20').Value2)
  if($isContractor){
    $payload.loan=0;$payload.loanMonthly=0
    $payload.gratuity=[math]::Round([double]$sheet.Range('C21').Value2)
    $payload.gratuityMonthly=[math]::Round([double]$sheet.Range('D21').Value2)
    $payload.benefits=[math]::Round([double]$sheet.Range('C23').Value2)
    $payload.benefitsMonthly=[math]::Round([double]$sheet.Range('D23').Value2)
    $payload.annualCtc=[math]::Round([double]$sheet.Range('C25').Value2)
    $payload.annualCtcMonthly=[math]::Round([double]$sheet.Range('D25').Value2)
  }else{
    $payload.loan=[math]::Round([double]$sheet.Range('C21').Value2)
    $payload.loanMonthly=[math]::Round([double]$sheet.Range('D21').Value2)
    $payload.gratuity=[math]::Round([double]$sheet.Range('C22').Value2)
    $payload.gratuityMonthly=[math]::Round([double]$sheet.Range('D22').Value2)
    $payload.benefits=[math]::Round([double]$sheet.Range('C24').Value2)
    $payload.benefitsMonthly=[math]::Round([double]$sheet.Range('D24').Value2)
    $payload.annualCtc=[math]::Round([double]$sheet.Range('C26').Value2)
    $payload.annualCtcMonthly=[math]::Round([double]$sheet.Range('D26').Value2)
  }
  if($exportAnnexure){
    $sheet.Range('C17').Interior.Pattern=-4142
    $sheet.PageSetup.PrintArea='$B$1:$D$25'
    $sheet.PageSetup.Orientation=1
    $sheet.PageSetup.PaperSize=9
    $sheet.PageSetup.Zoom=130
    $sheet.PageSetup.CenterHorizontally=$true
    $sheet.PageSetup.LeftMargin=$excel.InchesToPoints(0.35)
    $sheet.PageSetup.RightMargin=$excel.InchesToPoints(0.35)
    $sheet.PageSetup.TopMargin=$excel.InchesToPoints(0.45)
    $sheet.PageSetup.BottomMargin=$excel.InchesToPoints(0.45)
    $sheet.ExportAsFixedFormat(0,$annexurePath,0,$true,$false)
    $payload.salaryAnnexurePath=$annexurePath
  }
}

$excel=$null;$employeeBook=$null;$salaryBook=$null;$salaryCopy=$null
$records=New-Object System.Collections.Generic.List[object]
$errors=New-Object System.Collections.Generic.List[object]
try{
  $excel=New-Object -ComObject Excel.Application
  $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.ScreenUpdating=$false
  $employeeBook=$excel.Workbooks.Open([IO.Path]::GetFullPath($EmployeePath),0,$true)
  try{$source=$employeeBook.Worksheets.Item('Employees')}catch{$source=$employeeBook.Worksheets.Item(1)}
  $used=$source.UsedRange;$values=$used.Value2;$rowCount=$used.Rows.Count;$columnCount=$used.Columns.Count
  $headerRow=1;$foundHeader=$false
  for($rowIndex=1;$rowIndex-le[Math]::Min(25,$rowCount);$rowIndex++){
    for($columnIndex=1;$columnIndex-le$columnCount;$columnIndex++){
      if((Key ($values.GetValue($rowIndex,$columnIndex)))-in@('name','employeename','candidatename','employeeid','employeenumber')){
        $headerRow=$rowIndex;$foundHeader=$true;break
      }
    }
    if($foundHeader){break}
  }
  if(!$foundHeader){throw 'The employee Excel header row was not found. Use the Bulk Letter Import Template.'}
  $headers=@()
  for($columnIndex=1;$columnIndex-le$columnCount;$columnIndex++){$headers+=Key ($values.GetValue($headerRow,$columnIndex))}

  $letterType=[string]$base.type
  $salaryRequired=$letterType-in@('internal','contractor')
  if($letterType-eq'custom'-and(BoolValue $base.includeSalaryBreakup)-and(NumberValue $base.grossCtc)-gt0){$salaryRequired=$true}
  $isContractor=$letterType-eq'contractor'-or($letterType-eq'custom'-and[string]$base.salaryCategory-eq'contractor')
  if($salaryRequired){
    $master=if($isContractor){Join-Path $root 'templates\Aditi-Salary-Contractors.xlsm'}else{Join-Path $root 'templates\Aditi-Salary-Internal.xlsm'}
    $salaryCopy=Join-Path ([IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($ManifestPath))) 'salary-bulk-working.xlsm'
    Copy-Item -LiteralPath $master -Destination $salaryCopy -Force
    $salaryBook=$excel.Workbooks.Open($salaryCopy,0,$false)
    $salarySheet=$salaryBook.Worksheets.Item('Salary Breakup')
  }

  $sequence=0
  for($rowIndex=$headerRow+1;$rowIndex-le$rowCount;$rowIndex++){
    $row=@{}
    for($columnIndex=1;$columnIndex-le$columnCount;$columnIndex++){
      if($headers[$columnIndex-1]){$row[$headers[$columnIndex-1]]=$values.GetValue($rowIndex,$columnIndex)}
    }
    $name=Pick $row @('Employee Name','Candidate Name','Name')
    if([string]::IsNullOrWhiteSpace([string]$name)){continue}
    $sequence++
    try{
      $payload=New-Payload $base
      $payload.name=[string]$name
      Put $payload $row 'employeeId' @('Employee ID','Employee Number','Employee Code','Emp ID')
      Put $payload $row 'employeeEmail' @('Employee Email','Email','Official Email')
      Put $payload $row 'designation' @('Designation','Title','Job Title')
      Put $payload $row 'department' @('Department','Function')
      Put $payload $row 'currentDesignation' @('Current Designation','Existing Designation','Old Designation')
      Put $payload $row 'currentDepartment' @('Current Department','Existing Department','Old Department')
      Put $payload $row 'location' @('Location','Work Location')
      Put $payload $row 'address' @('Address','Current Address')
      Put $payload $row 'letterDate' @('Letter Date') 'date'
      Put $payload $row 'joiningDate' @('Joining Date','WO Start Date') 'date'
      Put $payload $row 'contractEndDate' @('Contract End Date','WO End Date') 'date'
      Put $payload $row 'acceptanceDate' @('Acceptance Date') 'date'
      Put $payload $row 'startDate' @('Start Date') 'date'
      Put $payload $row 'endDate' @('End Date','Last Working Day') 'date'
      Put $payload $row 'effectiveDate' @('Effective Date','Revision Date') 'date'
      Put $payload $row 'grossCtc' @('Gross CTC','Revised Gross CTC','New Gross CTC','Revised CTC','New CTC') 'number'
      Put $payload $row 'currentCtc' @('Current Gross CTC','Current CTC','Existing CTC') 'number'
      Put $payload $row 'incrementPercent' @('Increment %','Increment Percentage','Merit %','Merit Increase %') 'number'
      Put $payload $row 'revisedCtc' @('Revised Gross CTC','New Gross CTC','Revised CTC','New CTC') 'number'
      Put $payload $row 'pfAtActualBasic' @('PF at Actual Basic','PF Actual Basic') 'bool'
      Put $payload $row 'irtFrom' @('IRT From','IRT Start Date') 'date'
      Put $payload $row 'irtTo' @('IRT To','IRT End Date') 'date'
      Put $payload $row 'irtReason' @('IRT Reason','Reason')
      Put $payload $row 'irtManager' @('IRT Manager','Manager')
      Put $payload $row 'irtComments' @('IRT Comments','Comments')
      Put $payload $row 'irtSubject' @('IRT Subject','Subject')
      Put $payload $row 'signatory' @('Signatory','Signatory Name')
      Put $payload $row 'signatoryTitle' @('Signatory Title','Signatory Designation')

      if($salaryRequired){
        $appendAnnexure=$letterType-eq'custom'-and(BoolValue $base.appendSalaryAnnexure)
        $annexurePath=if($appendAnnexure){Join-Path $OutputDirectory ("{0:000}-{1}-{2}-salary-breakup.pdf"-f$sequence,(Safe $payload.employeeId),(Safe $payload.name))}else{''}
        Add-Salary $payload $salarySheet $excel $isContractor $appendAnnexure $annexurePath
      }
      $employeeId=Safe $payload.employeeId
      $employeeName=Safe $payload.name
      $typeLabel=Safe(([string]$letterType-replace'-',' '))
      $pdf=Join-Path $OutputDirectory ("{0:000}-{1}-{2}-{3}.pdf"-f$sequence,$employeeId,$employeeName,$typeLabel)
      $docx=if($letterType-eq'irt'){$pdf}else{Join-Path $OutputDirectory ("{0:000}-{1}-{2}-{3}.docx"-f$sequence,$employeeId,$employeeName,$typeLabel)}
      $json=Join-Path ([IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($ManifestPath))) ("employee-{0:000}.json"-f$sequence)
      $payload|ConvertTo-Json -Depth 6|Set-Content -LiteralPath $json -Encoding UTF8
      & (Join-Path $root 'generate-letter.ps1') -DataPath $json -OutputPath $docx
      $records.Add([ordered]@{row=$rowIndex;name=[string]$payload.name;employeeId=[string]$payload.employeeId;docxPath=$docx;pdfPath=$pdf;salaryAnnexurePath=[string]$payload.salaryAnnexurePath;dataPath=$json;fileName=[IO.Path]::GetFileName($pdf)})
    }catch{
      $errors.Add([ordered]@{row=$rowIndex;name=[string]$name;error=$_.Exception.Message})
    }
  }
  if($records.Count-eq0){
    $detail=if($errors.Count){$errors[0].error}else{'No employee names were found.'}
    throw "No letters were generated. $detail"
  }
  [ordered]@{generated=$records.Count;failed=$errors.Count;records=$records;errors=$errors}|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $ManifestPath -Encoding UTF8
}finally{
  if($employeeBook){$employeeBook.Close($false)}
  if($salaryBook){$salaryBook.Close($false)}
  if($excel){$excel.Quit()}
  if($salaryCopy-and(Test-Path -LiteralPath $salaryCopy)){Remove-Item -LiteralPath $salaryCopy -Force}
}
