param([Parameter(Mandatory=$true)][string]$TemplatePath,[Parameter(Mandatory=$true)][string]$EmployeePath,[ValidateSet('internal','contractor')][string]$Category='internal',[Parameter(Mandatory=$true)][string]$OutputZip)
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $MyInvocation.MyCommand.Path;$work=Split-Path -Parent ([IO.Path]::GetFullPath($OutputZip));$letters=Join-Path $work 'letters';New-Item -ItemType Directory -Force $letters|Out-Null
function Key($v){([string]$v).ToLowerInvariant()-replace'[^a-z0-9]',''}
function Pick($row,[string[]]$keys){foreach($k in $keys){$n=Key $k;if($row.ContainsKey($n)-and $null-ne$row[$n]-and[string]$row[$n]-ne''){return $row[$n]}};return $null}
function DateText($v){if($null -eq $v -or [string]$v -eq ''){return ''};if($v -is [double] -or $v -is [int]){return ([datetime]::FromOADate([double]$v).ToString('yyyy-MM-dd'))};try{return ((Get-Date $v).ToString('yyyy-MM-dd'))}catch{return ([string]$v)}}
function Safe($v){$s=([string]$v)-replace'[\\/:*?"<>|]','-';if(!$s){$s='Employee'};$s}
$excel=$null;$sourceBook=$null;$salaryBook=$null
try{
 $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false
 $sourceBook=$excel.Workbooks.Open([IO.Path]::GetFullPath($EmployeePath),0,$true);$source=$sourceBook.Worksheets.Item(1);$used=$source.UsedRange;$values=$used.Value2;$rows=$used.Rows.Count;$cols=$used.Columns.Count
 $headerRow=1;$foundHeader=$false;for($r=1;$r-le[Math]::Min(20,$rows);$r++){for($c=1;$c-le$cols;$c++){if((Key ($values.GetValue($r,$c)))-in@('name','employeename','employeeid','employeenumber')){$headerRow=$r;$foundHeader=$true;break}};if($foundHeader){break}}
 $headers=@();for($c=1;$c-le$cols;$c++){$headers+=Key ($values.GetValue($headerRow,$c))}
 $salaryMaster=Join-Path $root $(if($Category-eq'contractor'){'templates\Aditi-Salary-Contractors.xlsm'}else{'templates\Aditi-Salary-Internal.xlsm'});$salaryCopy=Join-Path $work 'salary-batch.xlsm';Copy-Item $salaryMaster $salaryCopy -Force;$salaryBook=$excel.Workbooks.Open($salaryCopy,0,$false);$salary=$salaryBook.Worksheets.Item('Salary Breakup')
 $count=0
 for($r=$headerRow+1;$r-le$rows;$r++){
  $row=@{};for($c=1;$c-le$cols;$c++){if($headers[$c-1]){$row[$headers[$c-1]]=$values.GetValue($r,$c)}}
  $name=Pick $row @('Employee Name','Name');if(!$name){continue};$employeeId=Pick $row @('Employee ID','Employee Number','Employee Code','Emp ID');$designation=Pick $row @('Designation','Title');$department=Pick $row @('Department','Function');$location=Pick $row @('Location','Work Location')
  $current=[double](Pick $row @('Current Gross CTC','Current CTC','Existing CTC','Current Salary'));$pct=[double](Pick $row @('Increment %','Increment Percentage','Merit %','Merit Increase %'));if($pct -gt 0 -and $pct -le 1){$pct=$pct*100};$revised=[double](Pick $row @('Revised Gross CTC','New Gross CTC','Revised CTC','New CTC'));if(!$revised -and $current){$revised=[math]::Round($current*(1+$pct/100),0)};if(!$revised){continue}
  $effective=DateText(Pick $row @('Effective Date','Increment Effective Date','Revision Date'));$letterDate=DateText(Pick $row @('Letter Date'));if(!$letterDate){$letterDate=(Get-Date).ToString('yyyy-MM-dd')}
  $pfActual=Pick $row @('PF at Actual Basic','PF Actual Basic');$pfChoice=if(([string]$pfActual).Trim().ToLowerInvariant()-in@('yes','y','true','1')){'Yes'}else{'No'}
  $salary.Range('C4').Value2=[string]$employeeId;$salary.Range('C5').Value2=[string]$name;$salary.Range('C6').Value2=[string]$designation;$salary.Range('C7').Value2=$pfChoice;$salary.Range('C17').Value2=$revised;$salary.Calculate()
  $basic=[math]::Round([double]$salary.Range('C10').Value2);$hra=[math]::Round([double]$salary.Range('C11').Value2);$bonus=[math]::Round([double]$salary.Range('C12').Value2);$special=[math]::Round([double]$salary.Range('C13').Value2);$grossSalary=[math]::Round([double]$salary.Range('C15').Value2);$pf=[math]::Round([double]$salary.Range('C16').Value2);$medical=[math]::Round([double]$salary.Range('C20').Value2)
  if($Category-eq'contractor'){$loan=0;$gratuity=[math]::Round([double]$salary.Range('C21').Value2);$benefits=[math]::Round([double]$salary.Range('C23').Value2);$overall=[math]::Round([double]$salary.Range('C25').Value2)}else{$loan=[math]::Round([double]$salary.Range('C21').Value2);$gratuity=[math]::Round([double]$salary.Range('C22').Value2);$benefits=[math]::Round([double]$salary.Range('C24').Value2);$overall=[math]::Round([double]$salary.Range('C26').Value2)}
  $payload=[ordered]@{type='custom';customTemplatePath=[IO.Path]::GetFullPath($TemplatePath);name=[string]$name;employeeId=[string]$employeeId;designation=[string]$designation;department=[string]$department;location=[string]$location;letterDate=$letterDate;effectiveDate=$effective;currentCtc=$current;incrementPercent=$pct;revisedCtc=$revised;grossCtc=$revised;annualCtc=$overall;basic=$basic;hra=$hra;bonus=$bonus;special=$special;grossSalary=$grossSalary;pf=$pf;medical=$medical;loan=$loan;gratuity=$gratuity;benefits=$benefits;signatory='Razia Khatoon';signatoryTitle='Associate Director - People & Culture'}
  $json=Join-Path $work "employee-$r.json";$docx=Join-Path $letters "$(Safe $employeeId)-$(Safe $name)-Merit-Increment.docx";$payload|ConvertTo-Json -Depth 4|Set-Content $json -Encoding UTF8
  & (Join-Path $root 'generate-letter.ps1') -DataPath $json -OutputPath $docx;$count++
 }
 if($count-eq0){throw 'No valid employee rows were found. Include Name and Revised Gross CTC, or Current CTC plus Increment %.'}
}finally{if($sourceBook){$sourceBook.Close($false)};if($salaryBook){$salaryBook.Close($false)};if($excel){$excel.Quit()}}
