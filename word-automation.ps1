if(-not('AditiWordNativeMethods'-as[type])){
 Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class AditiWordNativeMethods {
 [DllImport("user32.dll")]
 public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
}

function Remove-AditiOfficeBlockingMetadata([string]$path) {
 Add-Type -AssemblyName System.IO.Compression
 Add-Type -AssemblyName System.IO.Compression.FileSystem
 $archive=[IO.Compression.ZipFile]::Open($path,[IO.Compression.ZipArchiveMode]::Update)
 try{
  $labelPart=$archive.GetEntry('docMetadata/LabelInfo.xml')
  if($labelPart){$labelPart.Delete()}
  foreach($item in @(
   @('[Content_Types].xml','<Override\b[^>]*\bPartName="/docMetadata/LabelInfo\.xml"[^>]*/>'),
   @('_rels/.rels','<Relationship\b[^>]*\bType="[^"]*/classificationlabels"[^>]*/>'),
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

function Start-AditiWordSession {
 $wordPath=@(
  'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE',
  'C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE'
 )|Where-Object{Test-Path -LiteralPath $_}|Select-Object -First 1
 if(!$wordPath){throw 'Microsoft Word was not found on this computer.'}
 $existingProcesses=@(Get-Process WINWORD -ErrorAction SilentlyContinue)
 if($existingProcesses.Count){
  throw 'Please close Microsoft Word before generating letters. This keeps your open documents safe and makes PDF creation faster.'
 }
 $application=$null
 $process=$null
 try{
  # /a prevents workstation add-ins from delaying automation; /x starts a
  # dedicated instance; /automation makes it available immediately to COM.
  $launcher=Start-Process -FilePath $wordPath -ArgumentList @('/a','/x','/automation') -WindowStyle Minimized -PassThru
  if($env:ADITI_WORD_PID_FILE){[IO.File]::WriteAllText([IO.Path]::GetFullPath($env:ADITI_WORD_PID_FILE),[string]$launcher.Id)}
  for($attempt=1;$attempt-le80;$attempt++){
   try{
    $application=[Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
    if($application){break}
   }catch{}
   Start-Sleep -Milliseconds 250
  }
  if(!$application){throw 'Microsoft Word did not start its document service.'}
  $process=Get-Process -Id $launcher.Id -ErrorAction SilentlyContinue
  if(!$process){$process=Get-Process WINWORD -ErrorAction SilentlyContinue|Sort-Object StartTime|Select-Object -Last 1}
  if(!$process){throw 'Microsoft Word did not expose its background process.'}
  if($env:ADITI_WORD_PID_FILE){[IO.File]::WriteAllText([IO.Path]::GetFullPath($env:ADITI_WORD_PID_FILE),[string]$process.Id)}
  return [pscustomobject]@{Application=$application;Process=$process}
 }catch{
  try{if($application){$application.Quit()}}catch{}
  try{if($launcher-and!$launcher.HasExited){$launcher.Kill()}}catch{}
  throw
 }
}

function Stop-AditiWordSession($application,$process){
 try{if($application){$application.Quit()}}catch{}
 try{if($application){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)}}catch{}
 if($process){
  try{
   if(!$process.WaitForExit(5000)-and!$process.HasExited){$process.Kill()}
  }catch{}
 }
 if($env:ADITI_WORD_PID_FILE){Remove-Item -LiteralPath $env:ADITI_WORD_PID_FILE -Force -ErrorAction SilentlyContinue}
}
