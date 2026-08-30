$ErrorActionPreference = 'Stop'
$input = 'C:\ppt-creater\output\2026-ai-industry-report-native-template.pptx'
$out = 'C:\ppt-creater\catalog\native-ai-report-inspect.json'
$app = New-Object -ComObject PowerPoint.Application
$app.Visible = -1
$items = @()
try {
  $p = $app.Presentations.Open($input, $true, $false, $false)
  for ($i=1; $i -le $p.Slides.Count; $i++) {
    $s=$p.Slides.Item($i); $effects=@(); $texts=@()
    try { for($k=1;$k -le $s.TimeLine.MainSequence.Count;$k++){ $e=$s.TimeLine.MainSequence.Item($k); $effects += [ordered]@{index=$k; effect=$e.EffectType; trigger=$e.Timing.TriggerType; shape=$e.Shape.Name} } } catch {}
    for($j=1;$j -le $s.Shapes.Count;$j++){ $sh=$s.Shapes.Item($j); try {if($sh.HasTextFrame -and $sh.TextFrame.HasText){$texts += [ordered]@{index=$j;name=$sh.Name;text=$sh.TextFrame.TextRange.Text}}}catch{}}
    $items += [ordered]@{slide=$i; shapes=$s.Shapes.Count; animations=$s.TimeLine.MainSequence.Count; transition=$s.SlideShowTransition.EntryEffect; effects=$effects;texts=$texts}
  }
  $p.Close()
} finally {$app.Quit();[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null}
$items | ConvertTo-Json -Depth 7 | Set-Content -Path $out -Encoding utf8
