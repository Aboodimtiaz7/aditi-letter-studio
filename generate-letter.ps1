param([Parameter(Mandatory=$true)][string]$DataPath,[Parameter(Mandatory=$true)][string]$OutputPath)
$ErrorActionPreference='Stop'
$DataPath=[IO.Path]::GetFullPath($DataPath);$OutputPath=[IO.Path]::GetFullPath($OutputPath)
$d=Get-Content -Raw -LiteralPath $DataPath | ConvertFrom-Json
$root=Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $root 'word-automation.ps1')
$traceEnabled=$env:ADITI_LETTER_TRACE-eq'1';$tracePath="$OutputPath.trace.log"
function Trace([string]$message){if($traceEnabled){Add-Content -LiteralPath $tracePath -Value ("{0:HH:mm:ss.fff} {1}"-f(Get-Date),$message)}}
function ShortDate([string]$s){if(!$s){return ''};(Get-Date $s).ToString('dd-MMM-yyyy')}
function LongDate([string]$s){if(!$s){return ''};$x=Get-Date $s;$day=$x.Day;$suffix=if($day%100-in 11,12,13){'th'}elseif($day%10-eq 1){'st'}elseif($day%10-eq 2){'nd'}elseif($day%10-eq 3){'rd'}else{'th'};'{0:00}{1} {2:MMM yyyy}'-f $day,$suffix,$x}
function N($n){([double]$n).ToString('#,##0',[Globalization.CultureInfo]::GetCultureInfo('en-IN'))}
function Money($n){if($null-eq$n-or[string]::IsNullOrWhiteSpace([string]$n)){return ''};N $n}
function Words([long]$n){$o=@('','One','Two','Three','Four','Five','Six','Seven','Eight','Nine','Ten','Eleven','Twelve','Thirteen','Fourteen','Fifteen','Sixteen','Seventeen','Eighteen','Nineteen');$z=@('','','Twenty','Thirty','Forty','Fifty','Sixty','Seventy','Eighty','Ninety');function U([long]$v){$p=@();if($v-ge 100){$p+=$o[[math]::Floor($v/100)];$p+='Hundred';$v=$v%100};if($v-ge 20){$p+=$z[[math]::Floor($v/10)];if($v%10){$p+=$o[$v%10]}}elseif($v){$p+=$o[$v]};$p-join ' '};if($n-eq 0){return 'Zero'};$p=@();foreach($u in @(@(10000000,'Crore'),@(100000,'Lakh'),@(1000,'Thousand'))){if($n-ge $u[0]){$p+=(U ([math]::Floor($n/$u[0])));$p+=$u[1];$n=$n%$u[0]}};if($n){$p+=(U $n)};$p-join ' '}
function Find-Range($range,[string]$old,[string]$new){$f=$range.Find;$f.ClearFormatting();$f.Replacement.ClearFormatting();[void]$f.Execute($old,$false,$false,$false,$false,$false,$true,0,$false,$new,2)}
function Replace-All($doc,[string]$old,[string]$new){Find-Range $doc.Content $old $new}
function Replace-Header($doc,[string]$old,[string]$new){foreach($section in $doc.Sections){foreach($header in $section.Headers){Find-Range $header.Range $old $new;foreach($shape in $header.Shapes){if($shape.TextFrame.HasText){Find-Range $shape.TextFrame.TextRange $old $new}}}}}
function Replace-Everywhere($doc,[string]$old,[string]$new){
 Find-Range $doc.Content $old $new
 foreach($shape in $doc.Shapes){if($shape.TextFrame.HasText){Find-Range $shape.TextFrame.TextRange $old $new}}
 foreach($section in $doc.Sections){
  foreach($story in @($section.Headers)+@($section.Footers)){
   if($story.Exists){
    Find-Range $story.Range $old $new
    foreach($shape in $story.Shapes){if($shape.TextFrame.HasText){Find-Range $shape.TextFrame.TextRange $old $new}}
   }
  }
 }
}
function Replace-In-Paragraph($doc,[string]$needle,[string]$old,[string]$new){
 $search=$doc.Content.Duplicate
 $find=$search.Find
 $find.ClearFormatting()
 if($find.Execute($needle,$false,$false,$false,$false,$false,$true,0)){
  Find-Range $search.Paragraphs.Item(1).Range $old $new
 }
}
function Set-Paragraph-Alignment($doc,[string]$needle,[int]$alignment){
 $search=$doc.Content.Duplicate
 $find=$search.Find
 $find.ClearFormatting()
 if($find.Execute($needle,$false,$false,$false,$false,$false,$true,0)){
  $search.Paragraphs.Item(1).Alignment=$alignment
 }
}
function Clear-Highlights($doc){
 $doc.Content.HighlightColorIndex=0
 foreach($section in $doc.Sections){foreach($story in @($section.Headers)+@($section.Footers)){$story.Range.HighlightColorIndex=0}}
}
function Cell($table,[int]$row,[int]$col,[string]$value){$r=$table.Cell($row,$col).Range;$r.End=$r.End-1;$r.Text=$value}
function Text-Nodes($paragraph,$namespaceManager){return @($paragraph.SelectNodes('.//w:t',$namespaceManager))}
function Joined-Text($nodes){return (($nodes|ForEach-Object{$_.InnerText})-join'')}
function Preserve-Space($node){
 if($node.InnerText -match '^\s|\s$'){[void]$node.SetAttribute('space','http://www.w3.org/XML/1998/namespace','preserve')}
}
function Replace-In-Nodes($nodes,[string]$old,[string]$new){
 $nodes=@($nodes)
 if(!$nodes-or[string]::IsNullOrEmpty($old)){return $false}
 $original=Joined-Text $nodes
 $matches=New-Object System.Collections.Generic.List[int]
 $searchIndex=0
 while($searchIndex-le$original.Length-$old.Length){
  $found=$original.IndexOf($old,$searchIndex,[StringComparison]::Ordinal)
  if($found-lt0){break}
  [void]$matches.Add($found)
  $searchIndex=$found+$old.Length
 }
 if(!$matches.Count){return $false}
 for($matchNumber=$matches.Count-1;$matchNumber-ge0;$matchNumber--){
  $matchIndex=$matches[$matchNumber]
  $matchEnd=$matchIndex+$old.Length
  $position=0;$startIndex=-1;$endIndex=-1;$startOffset=0;$endOffset=0
  for($index=0;$index-lt$nodes.Count;$index++){
   $length=$nodes[$index].InnerText.Length
   $nodeEnd=$position+$length
   if($startIndex-lt0-and$matchIndex-ge$position-and$matchIndex-lt$nodeEnd){$startIndex=$index;$startOffset=$matchIndex-$position}
   if($matchEnd-gt$position-and$matchEnd-le$nodeEnd){$endIndex=$index;$endOffset=$matchEnd-$position;break}
   $position=$nodeEnd
  }
  if($startIndex-lt0-or$endIndex-lt0){continue}
  if($startIndex-eq$endIndex){
   $text=$nodes[$startIndex].InnerText
   $nodes[$startIndex].InnerText=$text.Substring(0,$startOffset)+$new+$text.Substring($endOffset)
   Preserve-Space $nodes[$startIndex]
  }else{
   $startText=$nodes[$startIndex].InnerText
   $endText=$nodes[$endIndex].InnerText
   $nodes[$startIndex].InnerText=$startText.Substring(0,$startOffset)+$new
   Preserve-Space $nodes[$startIndex]
   for($index=$startIndex+1;$index-lt$endIndex;$index++){$nodes[$index].InnerText=''}
   $nodes[$endIndex].InnerText=$endText.Substring($endOffset)
   Preserve-Space $nodes[$endIndex]
  }
 }
 return $true
}
function Load-Docx-Parts([string]$path){
 Add-Type -AssemblyName System.IO.Compression
 Add-Type -AssemblyName System.IO.Compression.FileSystem
 $archive=[IO.Compression.ZipFile]::Open($path,[IO.Compression.ZipArchiveMode]::Update)
 $parts=New-Object System.Collections.Generic.List[object]
 foreach($entry in @($archive.Entries)){
  if($entry.FullName -notmatch '^word/(document|header\d+|footer\d+|footnotes|endnotes)\.xml$'){continue}
  $reader=New-Object IO.StreamReader($entry.Open())
  try{$text=$reader.ReadToEnd()}finally{$reader.Dispose()}
  $xml=New-Object Xml.XmlDocument;$xml.PreserveWhitespace=$true;$xml.LoadXml($text)
  $namespaceManager=New-Object Xml.XmlNamespaceManager($xml.NameTable)
  $namespaceManager.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
  [void]$parts.Add([pscustomobject]@{Name=$entry.FullName;Xml=$xml;NamespaceManager=$namespaceManager;Changed=$false})
 }
 return [pscustomobject]@{Archive=$archive;Parts=$parts}
}
function Save-Docx-Parts($package){
 try{
  foreach($part in $package.Parts){
   if(!$part.Changed){continue}
   $entry=$package.Archive.GetEntry($part.Name)
   $entry.Delete()
   $replacement=$package.Archive.CreateEntry($part.Name,[IO.Compression.CompressionLevel]::Optimal)
   $settings=New-Object Xml.XmlWriterSettings
   $settings.Encoding=New-Object Text.UTF8Encoding($false)
   $settings.Indent=$false
   $stream=$replacement.Open()
   $writer=[Xml.XmlWriter]::Create($stream,$settings)
   try{$part.Xml.Save($writer)}finally{$writer.Dispose();$stream.Dispose()}
  }
 }finally{$package.Archive.Dispose()}
}
function Replace-Docx-Text($parts,[string]$old,[string]$new){
 foreach($part in $parts){
  foreach($paragraph in @($part.Xml.SelectNodes('//w:p',$part.NamespaceManager))){
   if(Replace-In-Nodes (Text-Nodes $paragraph $part.NamespaceManager) $old $new){$part.Changed=$true}
  }
 }
}
function Replace-Docx-Paragraph($parts,[string]$needle,[string]$old,[string]$new){
 foreach($part in $parts){
  foreach($paragraph in @($part.Xml.SelectNodes('//w:p',$part.NamespaceManager))){
   $nodes=Text-Nodes $paragraph $part.NamespaceManager
   if((Joined-Text $nodes).Contains($needle)){
    if(Replace-In-Nodes $nodes $old $new){$part.Changed=$true}
    return
   }
  }
 }
}
function Set-Docx-Cell($parts,[int]$tableNumber,[int]$rowNumber,[int]$columnNumber,[string]$value){
 $part=$parts|Where-Object{$_.Name -eq 'word/document.xml'}|Select-Object -First 1
 $tables=@($part.Xml.SelectNodes('//w:tbl',$part.NamespaceManager))
 if($tableNumber-gt$tables.Count){throw "The official template is missing table $tableNumber."}
 $rows=@($tables[$tableNumber-1].SelectNodes('./w:tr',$part.NamespaceManager))
 if($rowNumber-gt$rows.Count){throw "The official template is missing row $rowNumber in table $tableNumber."}
 $cells=@($rows[$rowNumber-1].SelectNodes('./w:tc',$part.NamespaceManager))
 if($columnNumber-gt$cells.Count){throw "The official template is missing column $columnNumber in table $tableNumber."}
 $nodes=@($cells[$columnNumber-1].SelectNodes('.//w:t',$part.NamespaceManager))
 if(!$nodes){throw "The official template has no editable text in table $tableNumber, row $rowNumber, column $columnNumber."}
 $nodes[0].InnerText=$value;Preserve-Space $nodes[0]
 for($index=1;$index-lt$nodes.Count;$index++){$nodes[$index].InnerText=''}
 $part.Changed=$true
}
function Clear-Docx-Highlights($parts){
 foreach($part in $parts){
  $highlights=@($part.Xml.SelectNodes('//w:highlight',$part.NamespaceManager))
  foreach($highlight in $highlights){[void]$highlight.ParentNode.RemoveChild($highlight);$part.Changed=$true}
 }
}
function Assert-No-TemplatePlaceholders($parts){
 $unresolved=New-Object System.Collections.Generic.HashSet[string]
 foreach($part in $parts){
  foreach($paragraph in @($part.Xml.SelectNodes('//w:p',$part.NamespaceManager))){
   $text=Joined-Text (Text-Nodes $paragraph $part.NamespaceManager)
   foreach($match in [regex]::Matches($text,'\{\{[A-Za-z][A-Za-z0-9]*\}\}')){[void]$unresolved.Add($match.Value)}
  }
 }
 if($unresolved.Count){throw "The uploaded official format contains unsupported field(s): $(([string[]]$unresolved)-join', '). Use the field names shown in Other Letter."}
}
function Fill-Docx([string]$path,$data,[bool]$uploadedTemplate){
 $package=Load-Docx-Parts $path
 try{
  $parts=$package.Parts
  if($data.type -eq 'custom'){
   $map=[ordered]@{
    '{{name}}'=$data.name;'{{employeeName}}'=$data.name;'{{pronoun}}'=$data.pronoun;'{{employeeId}}'=$data.employeeId;'{{employeeEmail}}'=$data.employeeEmail
    '{{designation}}'=$data.designation;'{{newDesignation}}'=$data.designation;'{{currentDesignation}}'=$data.currentDesignation
     '{{department}}'=$data.department;'{{newDepartment}}'=$data.department;'{{currentDepartment}}'=$data.currentDepartment
     '{{location}}'=$data.location;'{{address}}'=$data.address;'{{clientName}}'=$data.clientName
    '{{letterDate}}'=(ShortDate $data.letterDate);'{{joiningDate}}'=(ShortDate $data.joiningDate);'{{acceptanceDate}}'=(ShortDate $data.acceptanceDate)
    '{{startDate}}'=(ShortDate $data.startDate);'{{endDate}}'=(ShortDate $data.endDate);'{{lastWorkingDay}}'=(ShortDate $data.endDate)
     '{{contractEndDate}}'=(ShortDate $data.contractEndDate);'{{effectiveDate}}'=(ShortDate $data.effectiveDate)
     '{{tentativeStartDate}}'=(ShortDate $data.tentativeStartDate);'{{assignmentDuration}}'=$data.assignmentDuration
     '{{noticePeriod}}'=$data.noticePeriod;'{{probationNoticePeriod}}'=$data.probationNoticePeriod;'{{joiningBonus}}'=(Money $data.joiningBonus)
    '{{irtFrom}}'=(ShortDate $data.irtFrom);'{{irtTo}}'=(ShortDate $data.irtTo);'{{irtReason}}'=$data.irtReason;'{{irtManager}}'=$data.irtManager;'{{irtComments}}'=$data.irtComments
    '{{signatory}}'=$data.signatory;'{{signatoryName}}'=$data.signatory;'{{signatoryTitle}}'=$data.signatoryTitle
    '{{currentCtc}}'=(Money $data.currentCtc);'{{incrementPercent}}'=if($null-eq$data.incrementPercent){''}else{"$($data.incrementPercent)%"};'{{revisedCtc}}'=(Money $data.revisedCtc)
    '{{grossCtc}}'=(Money $data.grossCtc);'{{grossCtcWords}}'=if($data.grossCtc){Words ([long]$data.grossCtc)}else{''};'{{annualCtc}}'=(Money $data.annualCtc);'{{annualCtcWords}}'=if($data.annualCtc){Words ([long]$data.annualCtc)}else{''}
    '{{basic}}'=(Money $data.basic);'{{basicMonthly}}'=(Money $data.basicMonthly);'{{hra}}'=(Money $data.hra);'{{hraMonthly}}'=(Money $data.hraMonthly)
    '{{bonus}}'=(Money $data.bonus);'{{bonusMonthly}}'=(Money $data.bonusMonthly);'{{special}}'=(Money $data.special);'{{specialMonthly}}'=(Money $data.specialMonthly)
    '{{pf}}'=(Money $data.pf);'{{pfMonthly}}'=(Money $data.pfMonthly);'{{medical}}'=(Money $data.medical);'{{medicalMonthly}}'=(Money $data.medicalMonthly)
    '{{loan}}'=(Money $data.loan);'{{loanMonthly}}'=(Money $data.loanMonthly);'{{gratuity}}'=(Money $data.gratuity);'{{gratuityMonthly}}'=(Money $data.gratuityMonthly)
    '{{grossSalary}}'=(Money $data.grossSalary);'{{grossSalaryMonthly}}'=(Money $data.grossSalaryMonthly);'{{benefits}}'=(Money $data.benefits);'{{benefitsMonthly}}'=(Money $data.benefitsMonthly)
    '[Employee Name]'=$data.name;'[Employee ID]'=$data.employeeId;'[Employee Email]'=$data.employeeEmail;'[Designation]'=$data.designation;'[Current Designation]'=$data.currentDesignation
     '[Department]'=$data.department;'[Current Department]'=$data.currentDepartment;'[Location]'=$data.location;'[Address]'=$data.address;'[Client Name]'=$data.clientName
    '[Letter Date]'=(ShortDate $data.letterDate);'[Joining Date]'=(ShortDate $data.joiningDate);'[Acceptance Date]'=(ShortDate $data.acceptanceDate)
     '[Start Date]'=(ShortDate $data.startDate);'[End Date]'=(ShortDate $data.endDate);'[Last Working Day]'=(ShortDate $data.endDate);'[Contract End Date]'=(ShortDate $data.contractEndDate);'[Effective Date]'=(ShortDate $data.effectiveDate);'[Tentative Start Date]'=(ShortDate $data.tentativeStartDate)
     '[Assignment Duration]'=$data.assignmentDuration;'[Notice Period]'=$data.noticePeriod;'[Probation Notice Period]'=$data.probationNoticePeriod;'[Joining Bonus]'=(Money $data.joiningBonus)
    '[Signatory Name]'=$data.signatory;'[Signatory Title]'=$data.signatoryTitle
   }
    foreach($key in $map.Keys){Replace-Docx-Text $parts $key ([string]$map[$key])}
    Assert-No-TemplatePlaceholders $parts
   }elseif([string]$data.type -like 'india-*'){
    $duration=if([string]::IsNullOrWhiteSpace([string]$data.assignmentDuration)){'6 months'}else{[string]$data.assignmentDuration}
    $notice=if([string]::IsNullOrWhiteSpace([string]$data.noticePeriod)){'30 days'}else{[string]$data.noticePeriod}
    if($uploadedTemplate){
     $map=[ordered]@{
      '{{name}}'=$data.name;'{{employeeName}}'=$data.name;'{{employeeId}}'=$data.employeeId;'{{designation}}'=$data.designation
      '{{clientName}}'=$data.clientName;'{{location}}'=$data.location;'{{address}}'=$data.address
      '{{letterDate}}'=(LongDate $data.letterDate);'{{joiningDate}}'=(LongDate $data.joiningDate);'{{tentativeStartDate}}'=(LongDate $data.tentativeStartDate)
      '{{contractEndDate}}'=(LongDate $data.contractEndDate);'{{assignmentDuration}}'=$duration;'{{noticePeriod}}'=$notice
      '{{grossCtc}}'=(Money $data.grossCtc);'{{joiningBonus}}'=(Money $data.joiningBonus);'{{acceptanceDate}}'=(LongDate $data.acceptanceDate)
      '{{signatoryName}}'=$data.signatory;'{{signatoryTitle}}'=$data.signatoryTitle
     }
     foreach($key in $map.Keys){Replace-Docx-Text $parts $key ([string]$map[$key])}
    }
    if($data.type -eq 'india-intent'){
     Replace-Docx-Paragraph $parts 'Address:' 'Date:' ("Date: $(LongDate $data.letterDate)")
     Replace-Docx-Paragraph $parts 'Address:' 'Name:' ("Name: $($data.name)")
     Replace-Docx-Paragraph $parts 'Address:' 'Address:' ("Address: $($data.address)")
     Replace-Docx-Text $parts '[Candidate Name]' ([string]$data.name)
     Replace-Docx-Text $parts '[Designation]' ([string]$data.designation)
     Replace-Docx-Text $parts '[Tentative Start Date]' (LongDate $data.tentativeStartDate)
     Replace-Docx-Text $parts 'XX Months' $duration
     Replace-Docx-Text $parts '[Amount]' (Money $data.grossCtc)
     Replace-Docx-Text $parts '[Date]' (LongDate $data.acceptanceDate)
    }elseif($data.type -eq 'india-appointment'){
     Replace-Docx-Paragraph $parts 'EmployeeName:' 'Date:' ("Date: $(LongDate $data.letterDate)")
     Replace-Docx-Paragraph $parts 'EmployeeName:' 'EmployeeName:' ("EmployeeName: $($data.name)")
     Replace-Docx-Paragraph $parts 'EmployeeName:' 'Address:' ("Address: $($data.address)")
     Replace-Docx-Text $parts '[Employee Name]' ([string]$data.name)
     Replace-Docx-Text $parts '[Designation]' ([string]$data.designation)
     Replace-Docx-Text $parts '[Client Name]' ([string]$data.clientName)
     Replace-Docx-Text $parts '[Location]' ([string]$data.location)
     Replace-Docx-Text $parts '[WO START DATE]' (LongDate $data.joiningDate)
     Replace-Docx-Text $parts '[WO END DATE]' (LongDate $data.contractEndDate)
     Replace-Docx-Text $parts '[XX] months' $duration
    }else{
     Replace-Docx-Text $parts 'Candidate Name' ([string]$data.name)
     Replace-Docx-Text $parts 'Job Title' ([string]$data.designation)
     Replace-Docx-Text $parts '25000' (Money $data.joiningBonus)
     Replace-Docx-Text $parts '[initial WO duration]' $duration
    }
    Replace-Docx-Text $parts 'giving 30 days' ("giving $notice")
    Replace-Docx-Text $parts 'thirty (30) days' $notice
    if($data.type -in @('india-intent','india-appointment')){
     Replace-Docx-Text $parts 'AuthorizedSignatory' 'Authorized Signatory'
     Replace-Docx-Text $parts 'Name:___________________' ("Name: $($data.signatory)")
     Replace-Docx-Text $parts 'Designation: ___________________' ("Designation: $($data.signatoryTitle)")
     Replace-Docx-Text $parts 'Name:________________________' ("Name: $($data.signatory)")
     Replace-Docx-Text $parts 'Designation:__________________' ("Designation: $($data.signatoryTitle)")
     Replace-Docx-Text $parts 'Date: _________________________' ("Date: $(LongDate $data.letterDate)")
    }
   }elseif($data.type -eq 'experience'){
   $map=[ordered]@{'18th Jun 2026'=(LongDate $data.letterDate);'Mr. Aaditya Shahapurkar'=("$($data.pronoun) $($data.name)");'ASIPL3875'=$data.employeeId;'05 May 2025'=((Get-Date $data.startDate).ToString('dd MMM yyyy'));'14 Jul 2025'=((Get-Date $data.endDate).ToString('dd MMM yyyy'));'Technical Recruiter'=$data.designation;'Razia Khatoon'=$data.signatory;'Associate Director - People & Culture'=$data.signatoryTitle}
   foreach($key in $map.Keys){Replace-Docx-Text $parts $key ([string]$map[$key])}
  }else{
   Replace-Docx-Text $parts 'Dear Name,' ("Dear $($data.name),")
   Replace-Docx-Text $parts 'position of Title' ("position of $($data.designation)")
   $bonus=if($null-eq$data.bonus){0}else{[double]$data.bonus};$loan=if($null-eq$data.loan){0}else{[double]$data.loan};$gross=if($data.grossSalary){[double]$data.grossSalary}else{[double]$data.basic+[double]$data.hra+$bonus+[double]$data.special};$grossCtc=if($data.grossCtc){[double]$data.grossCtc}else{$gross+[double]$data.pf};$benefits=if($data.benefits){[double]$data.benefits}else{[double]$data.medical+$loan+[double]$data.gratuity}
    if($data.type -eq 'internal'){
    Replace-Docx-Paragraph $parts 'returning a copy of both the documents' '04th May 2026' (LongDate $data.acceptanceDate)
    $map=[ordered]@{'30th Apr 2026'=(LongDate $data.letterDate);'Bangalore'=$data.location;'04th May 2026'=(LongDate $data.joiningDate);'16,00,000'=(N $grossCtc);'18,00,000'=(N $data.annualCtc);'5,40,000'=(N $data.basic);'8,60,000'=(N ([double]$data.hra+$bonus+[double]$data.special+[double]$data.pf));'(Rupees Sixteen Lakh Only)'=("(Rupees $(Words ([long]$grossCtc)) Only)");'(Rupees Eighteen lakh Only)'=("(Rupees $(Words ([long]$data.annualCtc)) Only)");'Ramna Vadavalli'=$data.signatory;'India Country Head'=$data.signatoryTitle}
     foreach($key in $map.Keys){Replace-Docx-Text $parts $key ([string]$map[$key])}
     $notice=if([string]::IsNullOrWhiteSpace([string]$data.noticePeriod)){'3 Months'}else{[string]$data.noticePeriod}
     $probationNotice=if([string]::IsNullOrWhiteSpace([string]$data.probationNoticePeriod)){'15 days'}else{[string]$data.probationNoticePeriod}
     Replace-Docx-Text $parts 'Notice period to be served is as 3 Months.' ("Notice period to be served is as $notice.")
     Replace-Docx-Text $parts 'notice of 15 days' ("notice of $probationNotice")
     Replace-Docx-Text $parts 'i.e., 15 days.' ("i.e., $probationNotice.")
    Set-Docx-Cell $parts 1 2 2 ([string]$data.name);Set-Docx-Cell $parts 1 3 2 ([string]$data.designation)
    Set-Docx-Cell $parts 2 15 2 '';Set-Docx-Cell $parts 2 15 3 '';Set-Docx-Cell $parts 2 17 2 '';Set-Docx-Cell $parts 2 17 3 ''
    $rows=@(@(2,$data.basic),@(3,$data.hra),@(4,$bonus),@(5,$data.special),@(7,$gross),@(8,$data.pf),@(9,$grossCtc),@(12,$data.medical),@(13,$loan),@(14,$data.gratuity),@(16,$benefits),@(18,$data.annualCtc))
    }else{
    $map=[ordered]@{'current location of work will be Location'=("current location of work will be $($data.location)");'7-Jul-2026'=(ShortDate $data.letterDate);'06th Jul 2026'=(LongDate $data.joiningDate);'05th Jan 2027'=(LongDate $data.contractEndDate);'10,00,000'=(N $grossCtc);'18,00,000'=(N $data.annualCtc);'(Rupees Ten Lakh only)'=("(Rupees $(Words ([long]$grossCtc)) Only)");'(Rupees Eighteen Lakh only)'=("(Rupees $(Words ([long]$data.annualCtc)) Only)")}
     foreach($key in $map.Keys){Replace-Docx-Text $parts $key ([string]$map[$key])}
     $notice=if([string]::IsNullOrWhiteSpace([string]$data.noticePeriod)){'30 days'}else{[string]$data.noticePeriod}
     Replace-Docx-Text $parts 'giving 30 days' ("giving $notice")
     Replace-Docx-Text $parts 'thirty (30) days' $notice
    Set-Docx-Cell $parts 1 1 2 ([string]$data.name);Set-Docx-Cell $parts 1 2 2 ([string]$data.designation)
    Set-Docx-Cell $parts 2 15 2 '';Set-Docx-Cell $parts 2 15 3 ''
    $rows=@(@(2,$data.basic),@(3,$data.hra),@(4,$data.special),@(6,$gross),@(7,$data.pf),@(8,$grossCtc),@(11,$data.medical),@(12,$data.gratuity),@(14,$benefits),@(16,$data.annualCtc))
   }
   foreach($row in $rows){Set-Docx-Cell $parts 2 $row[0] 2 (N $row[1]);Set-Docx-Cell $parts 2 $row[0] 3 (N ([math]::Round([double]$row[1]/12)))}
  }
   if($data.manualReplacements){
    foreach($edit in @($data.manualReplacements)){
     $source=[string]$edit.source
     $replacement=[string]$edit.replacement
     if(-not[string]::IsNullOrWhiteSpace($source)){Replace-Docx-Text $parts $source $replacement}
    }
   }
   if(($data.type -in @('internal','contractor')) -or ([string]$data.type -like 'india-*')){Clear-Docx-Highlights $parts}
 }finally{Save-Docx-Parts $package}
}
function Remove-Attached-Template([string]$path){
 Add-Type -AssemblyName System.IO.Compression
 Add-Type -AssemblyName System.IO.Compression.FileSystem
 $archive=[IO.Compression.ZipFile]::Open($path,[IO.Compression.ZipArchiveMode]::Update)
 try{
  foreach($item in @(
   @('word/settings.xml','<w:attachedTemplate\b[^>]*/>'),
   @('word/_rels/settings.xml.rels','<Relationship\b[^>]*\bType="[^"]*/attachedTemplate"[^>]*/>')
  )){
   $entry=$archive.GetEntry($item[0])
   if(!$entry){continue}
   $reader=New-Object IO.StreamReader($entry.Open())
   try{$text=$reader.ReadToEnd()}finally{$reader.Dispose()}
   $updated=[regex]::Replace($text,$item[1],'',[Text.RegularExpressions.RegexOptions]::IgnoreCase)
   if($updated-ne$text){
    $entry.Delete()
    $replacement=$archive.CreateEntry($item[0],[IO.Compression.CompressionLevel]::Optimal)
    $writer=New-Object IO.StreamWriter($replacement.Open(),(New-Object Text.UTF8Encoding($false)))
    try{$writer.Write($updated)}finally{$writer.Dispose()}
   }
  }
 }finally{$archive.Dispose()}
}
$uploaded=$d.customTemplatePath
if($uploaded -and [IO.Path]::GetExtension([string]$uploaded) -eq '.pdf'){Copy-Item -LiteralPath $uploaded -Destination $OutputPath -Force;exit 0}
if($d.type-eq'irt'){Copy-Item -LiteralPath $(if($uploaded){$uploaded}else{Join-Path $root 'templates\IRT.pdf'}) -Destination $OutputPath -Force;exit 0}
$template=if($uploaded){$uploaded}else{switch($d.type){'internal'{Join-Path $root 'templates\Internal-Appointment-Letter.docx'}'contractor'{Join-Path $root 'templates\Contractor-Appointment-Letter.docx'}'experience'{Join-Path $root 'templates\Relieving-Experience-Letter.docx'}'india-intent'{Join-Path $root 'templates\India-Onboarding-Intent-of-Offer.docx'}'india-appointment'{Join-Path $root 'templates\India-Onboarding-Contractor-Appointment.docx'}'india-joining-bonus'{Join-Path $root 'templates\India-Onboarding-Joining-Bonus.docx'}'custom'{$d.customTemplatePath}default{throw 'Unknown letter type'}}}
$workingDirectory=Join-Path ([IO.Path]::GetTempPath()) ("Aditi-Letter-"+[guid]::NewGuid().ToString('N'))
$docx=Join-Path $workingDirectory 'letter.docx'
$tempPdf=Join-Path $workingDirectory 'letter.pdf'
$word=$null;$wordProcess=$null;$doc=$null;$sourceDocument=$null;$ownsWord=$false
try{
 New-Item -ItemType Directory -Path $workingDirectory -Force|Out-Null
 if([IO.Path]::GetExtension([string]$template) -eq '.odt'){
  if($global:AditiWordApplication){$word=$global:AditiWordApplication}else{$wordSession=Start-AditiWordSession;$word=$wordSession.Application;$wordProcess=$wordSession.Process;$ownsWord=$true}
  $word.Visible=$false;$word.ScreenUpdating=$false;$word.DisplayAlerts=0;$word.AutomationSecurity=3;$word.Options.ConfirmConversions=$false;$word.Options.UpdateLinksAtOpen=$false;$word.Options.CheckSpellingAsYouType=$false;$word.Options.CheckGrammarAsYouType=$false;$word.Options.BackgroundSave=$false;$word.Options.PrintBackground=$false
  $odt=Join-Path $workingDirectory 'source.odt'
  Copy-Item -LiteralPath $template -Destination $odt -Force
  Unblock-File -LiteralPath $odt -ErrorAction SilentlyContinue
  $sourceDocument=$word.Documents.Open($odt,$false,$true,$false,'','',$false,'','',18)
  $sourceDocument.SaveAs2($docx,16)
  $sourceDocument.Close($false)
  $sourceDocument=$null
 }else{Copy-Item -LiteralPath $template -Destination $docx -Force;Unblock-File -LiteralPath $docx -ErrorAction SilentlyContinue;Remove-Attached-Template $docx;Remove-AditiOfficeBlockingMetadata $docx}
 Trace 'Working copy prepared'
 Fill-Docx $docx $d ([bool]$uploaded)
 Trace 'DOCX fields filled directly'
 if([IO.Path]::GetExtension($OutputPath) -eq '.docx'){
  Copy-Item -LiteralPath $docx -Destination $OutputPath -Force
  Trace 'Filled DOCX exported'
  return
 }
 if(!$word){
  if($global:AditiWordApplication){$word=$global:AditiWordApplication}else{$wordSession=Start-AditiWordSession;$word=$wordSession.Application;$wordProcess=$wordSession.Process;$ownsWord=$true}
  $word.Visible=$false;$word.ScreenUpdating=$false;$word.DisplayAlerts=0;$word.AutomationSecurity=3;$word.Options.ConfirmConversions=$false;$word.Options.UpdateLinksAtOpen=$false;$word.Options.CheckSpellingAsYouType=$false;$word.Options.CheckGrammarAsYouType=$false;$word.Options.BackgroundSave=$false;$word.Options.PrintBackground=$false
 }
 $doc=$word.Documents.Open($docx,$false,$true,$false)
 Trace 'Filled document opened read-only'
 $doc.ExportAsFixedFormat($tempPdf,17)
 Copy-Item -LiteralPath $tempPdf -Destination $OutputPath -Force
 Trace 'PDF exported'
 return
 $doc=$word.Documents.Open($docx,$false,$false,$false)
 Trace 'Document opened'
 if($d.type-eq'custom'){
  $map=@{
   '{{name}}'=$d.name;'{{employeeName}}'=$d.name;'{{employeeId}}'=$d.employeeId;'{{designation}}'=$d.designation;'{{department}}'=$d.department;'{{location}}'=$d.location;'{{letterDate}}'=(ShortDate $d.letterDate);'{{joiningDate}}'=(ShortDate $d.joiningDate);'{{startDate}}'=(ShortDate $d.startDate);'{{endDate}}'=(ShortDate $d.endDate);'{{contractEndDate}}'=(ShortDate $d.contractEndDate);'{{effectiveDate}}'=(ShortDate $d.effectiveDate);'{{signatory}}'=$d.signatory;'{{signatoryTitle}}'=$d.signatoryTitle;'{{currentCtc}}'=(N $d.currentCtc);'{{incrementPercent}}'=("$($d.incrementPercent)%");'{{revisedCtc}}'=(N $d.revisedCtc);'{{grossCtc}}'=(N $d.grossCtc);'{{annualCtc}}'=(N $d.annualCtc);'{{basic}}'=(N $d.basic);'{{hra}}'=(N $d.hra);'{{bonus}}'=(N $d.bonus);'{{special}}'=(N $d.special);'{{pf}}'=(N $d.pf);'{{medical}}'=(N $d.medical);'{{loan}}'=(N $d.loan);'{{gratuity}}'=(N $d.gratuity);'{{grossSalary}}'=(N $d.grossSalary);'{{benefits}}'=(N $d.benefits);'[Employee Name]'=$d.name;'[Employee ID]'=$d.employeeId;'[Designation]'=$d.designation;'[Department]'=$d.department;'[Location]'=$d.location;'[Letter Date]'=(ShortDate $d.letterDate);'[Joining Date]'=(ShortDate $d.joiningDate);'[Start Date]'=(ShortDate $d.startDate);'[End Date]'=(ShortDate $d.endDate);'[Effective Date]'=(ShortDate $d.effectiveDate);'[Signatory Name]'=$d.signatory
  }
  foreach($key in $map.Keys){Replace-All $doc $key ([string]$map[$key]);Replace-Header $doc $key ([string]$map[$key])}
 }elseif($d.type-eq'experience'){
  Replace-All $doc '18th Jun 2026' (LongDate $d.letterDate);Replace-All $doc 'Aaditya Shahapurkar' $d.name;Replace-All $doc 'ASIPL3875' $d.employeeId;Replace-All $doc '05 May 2025' ((Get-Date $d.startDate).ToString('dd MMM yyyy'));Replace-All $doc '14 Jul 2025' ((Get-Date $d.endDate).ToString('dd MMM yyyy'));Replace-All $doc 'Technical Recruiter' $d.designation;Replace-All $doc 'Razia Khatoon' $d.signatory;Replace-All $doc 'Associate Director - People & Culture' $d.signatoryTitle
 }else{
  Replace-All $doc 'Dear Name,' ("Dear $($d.name),");Replace-All $doc 'position of Title' ("position of $($d.designation)")
  $bonus=if($null-eq$d.bonus){0}else{[double]$d.bonus};$loan=if($null-eq$d.loan){0}else{[double]$d.loan};$gross=if($d.grossSalary){[double]$d.grossSalary}else{[double]$d.basic+[double]$d.hra+$bonus+[double]$d.special};$grossCtc=if($d.grossCtc){[double]$d.grossCtc}else{$gross+[double]$d.pf};$benefits=if($d.benefits){[double]$d.benefits}else{[double]$d.medical+$loan+[double]$d.gratuity}
  if($d.type-eq'internal'){
   Replace-Header $doc '30th Apr 2026' (LongDate $d.letterDate);Replace-All $doc 'Bangalore' $d.location;Replace-All $doc '04th May 2026' (LongDate $d.joiningDate);Replace-All $doc '16,00,000' (N $grossCtc);Replace-All $doc '18,00,000' (N $d.annualCtc);Replace-All $doc '5,40,000' (N $d.basic);Replace-All $doc '8,60,000' (N ([double]$d.hra+$bonus+[double]$d.special+[double]$d.pf));Replace-All $doc '(Rupees Sixteen Lakh Only)' ("(Rupees $(Words ([long]$grossCtc)) Only)");Replace-All $doc '(Rupees Eighteen lakh Only)' ("(Rupees $(Words ([long]$d.annualCtc)) Only)")
   $s=$doc.Tables.Item(1);Cell $s 2 2 $d.name;Cell $s 3 2 $d.designation;$t=$doc.Tables.Item(2);Cell $t 15 2 '';Cell $t 15 3 '';Cell $t 17 2 '';Cell $t 17 3 '';$rows=@(@(2,$d.basic),@(3,$d.hra),@(4,$bonus),@(5,$d.special),@(7,$gross),@(8,$d.pf),@(9,$grossCtc),@(12,$d.medical),@(13,$loan),@(14,$d.gratuity),@(16,$benefits),@(18,$d.annualCtc))
  }else{
   Replace-All $doc 'current location of work will be Location' ("current location of work will be $($d.location)");Replace-All $doc '7-Jul-2026' (ShortDate $d.letterDate);Replace-All $doc '06th Jul 2026' (LongDate $d.joiningDate);Replace-All $doc '05th Jan 2027' (LongDate $d.contractEndDate);Replace-All $doc '10,00,000' (N $grossCtc);Replace-All $doc '18,00,000' (N $d.annualCtc);Replace-All $doc '(Rupees Ten Lakh only)' ("(Rupees $(Words ([long]$grossCtc)) Only)");Replace-All $doc '(Rupees Eighteen Lakh only)' ("(Rupees $(Words ([long]$d.annualCtc)) Only)")
   $s=$doc.Tables.Item(1);Cell $s 1 2 $d.name;Cell $s 2 2 $d.designation;$t=$doc.Tables.Item(2);Cell $t 15 2 '';Cell $t 15 3 '';$rows=@(@(2,$d.basic),@(3,$d.hra),@(4,$d.special),@(6,$gross),@(7,$d.pf),@(8,$grossCtc),@(11,$d.medical),@(12,$d.gratuity),@(14,$benefits),@(16,$d.annualCtc))
  }
 foreach($x in $rows){Cell $t $x[0] 2 (N $x[1]);Cell $t $x[0] 3 (N ([math]::Round([double]$x[1]/12)))}
 }
 if($d.type -in @('internal','contractor')){Clear-Highlights $doc}
 Trace 'Highlights cleared'
 $doc.Save()
 Trace 'Working document saved'
 $doc.Close($false);$doc=$null
 Trace 'Editing document closed'
 $doc=$word.Documents.Open($docx,$false,$true,$false)
 Trace 'Saved document reopened for PDF'
 $doc.ExportAsFixedFormat($tempPdf,17);Copy-Item -LiteralPath $tempPdf -Destination $OutputPath -Force
 Trace 'PDF exported'
}finally{if($doc){$doc.Close($false)};if($sourceDocument){$sourceDocument.Close($false)};if($word -and $ownsWord){Stop-AditiWordSession $word $wordProcess};if(Test-Path $workingDirectory){Remove-Item -LiteralPath $workingDirectory -Recurse -Force}}
