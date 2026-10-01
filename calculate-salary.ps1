param([Parameter(Mandatory=$true)][string]$DataPath,[Parameter(Mandatory=$true)][string]$OutputPath)
$ErrorActionPreference='Stop'
$d=Get-Content -Raw -LiteralPath $DataPath | ConvertFrom-Json
$root=Split-Path -Parent $MyInvocation.MyCommand.Path
$isContractor=$d.type -eq 'contractor' -or ($d.type -eq 'custom' -and [string]$d.salaryCategory -eq 'contractor')
$master=Join-Path $root $(if($isContractor){'templates\Aditi-Salary-Contractors.xlsm'}else{'templates\Aditi-Salary-Internal.xlsm'})
$temp=Join-Path ([IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($OutputPath))) 'salary-working.xlsm'
Copy-Item -LiteralPath $master -Destination $temp -Force
$excel=$null;$book=$null
try{
  $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.ScreenUpdating=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
  $book=$excel.Workbooks.Open($temp,0,$false);$sheet=$book.Worksheets.Item('Salary Breakup')
  $sheet.Range('C4').Value2=[string]$d.employeeId;$sheet.Range('C5').Value2=[string]$d.name;$sheet.Range('C6').Value2=[string]$d.designation
  $sheet.Range('C7').Value2=if($d.pfAtActualBasic){'Yes'}else{'No'}
  $sheet.Range('C17').Value2=[double]$d.grossCtc
  $sheet.Calculate()
  $result=[ordered]@{
    grossCtc=[math]::Round([double]$sheet.Range('C17').Value2);basic=[math]::Round([double]$sheet.Range('C10').Value2);basicMonthly=[math]::Round([double]$sheet.Range('D10').Value2)
    hra=[math]::Round([double]$sheet.Range('C11').Value2);hraMonthly=[math]::Round([double]$sheet.Range('D11').Value2);bonus=[math]::Round([double]$sheet.Range('C12').Value2);bonusMonthly=[math]::Round([double]$sheet.Range('D12').Value2)
    special=[math]::Round([double]$sheet.Range('C13').Value2);specialMonthly=[math]::Round([double]$sheet.Range('D13').Value2);grossSalary=[math]::Round([double]$sheet.Range('C15').Value2);grossSalaryMonthly=[math]::Round([double]$sheet.Range('D15').Value2)
    pf=[math]::Round([double]$sheet.Range('C16').Value2);pfMonthly=[math]::Round([double]$sheet.Range('D16').Value2);medical=[math]::Round([double]$sheet.Range('C20').Value2);medicalMonthly=[math]::Round([double]$sheet.Range('D20').Value2)
  }
  if($isContractor){$result.loan=0;$result.loanMonthly=0;$result.gratuity=[math]::Round([double]$sheet.Range('C21').Value2);$result.gratuityMonthly=[math]::Round([double]$sheet.Range('D21').Value2);$result.benefits=[math]::Round([double]$sheet.Range('C23').Value2);$result.benefitsMonthly=[math]::Round([double]$sheet.Range('D23').Value2);$result.annualCtc=[math]::Round([double]$sheet.Range('C25').Value2);$result.annualCtcMonthly=[math]::Round([double]$sheet.Range('D25').Value2)}
  else{$result.loan=[math]::Round([double]$sheet.Range('C21').Value2);$result.loanMonthly=[math]::Round([double]$sheet.Range('D21').Value2);$result.gratuity=[math]::Round([double]$sheet.Range('C22').Value2);$result.gratuityMonthly=[math]::Round([double]$sheet.Range('D22').Value2);$result.benefits=[math]::Round([double]$sheet.Range('C24').Value2);$result.benefitsMonthly=[math]::Round([double]$sheet.Range('D24').Value2);$result.annualCtc=[math]::Round([double]$sheet.Range('C26').Value2);$result.annualCtcMonthly=[math]::Round([double]$sheet.Range('D26').Value2)}
  if($d.type -eq 'custom' -and $d.appendSalaryAnnexure){
    $sheet.Range('C17').Interior.Pattern=-4142
    $sheet.PageSetup.PrintArea=if($isContractor){'$B$1:$D$25'}else{'$B$1:$D$26'}
    $sheet.PageSetup.Orientation=1;$sheet.PageSetup.PaperSize=9;$sheet.PageSetup.Zoom=130;$sheet.PageSetup.CenterHorizontally=$true
    $sheet.PageSetup.LeftMargin=$excel.InchesToPoints(0.35);$sheet.PageSetup.RightMargin=$excel.InchesToPoints(0.35);$sheet.PageSetup.TopMargin=$excel.InchesToPoints(0.45);$sheet.PageSetup.BottomMargin=$excel.InchesToPoints(0.45)
    $annexure=Join-Path ([IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($OutputPath))) 'salary-annexure.pdf'
    $sheet.ExportAsFixedFormat(0,$annexure,0,$true,$false)
    $result.salaryAnnexurePath=$annexure
  }
  $result|ConvertTo-Json -Depth 3|Set-Content -LiteralPath $OutputPath -Encoding UTF8
}finally{if($book){$book.Close($false)};if($excel){$excel.Quit()};if(Test-Path $temp){Remove-Item -LiteralPath $temp -Force}}
