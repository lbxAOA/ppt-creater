$ErrorActionPreference = 'Stop'
$input = 'C:\PPT模板\01_应用场景\02_汇报与答辩\模板\101_16比9比例.pptx'
$out = 'C:\ppt-creater\catalog\template-101-powerpoint-inspect.json'
$app = New-Object -ComObject PowerPoint.Application
$app.Visible = -1
$items = @()
try {
  $p = $app.Presentations.Open($input, $true, $false, $false)
  for ($i=1; $i -le $p.Slides.Count; $i++) {
    $s=$p.Slides.Item($i)
    $shapes=@()
    for ($j=1; $j -le $s.Shapes.Count; $j++) {
      $sh=$s.Shapes.Item($j)
      $text=''
      try { if ($sh.HasTextFrame -and $sh.TextFrame.HasText) { $text=$sh.TextFrame.TextRange.Text } } catch {}
      $shapes += [ordered]@{index=$j; name=$sh.Name; type=$sh.Type; text=$text.Substring(0,[Math]::Min($text.Length,250)); left=[math]::Round($sh.Left); top=[math]::Round($sh.Top); width=[math]::Round($sh.Width); height=[math]::Round($sh.Height)}
    }
    $effects=@()
    try { for($k=1;$k -le $s.TimeLine.MainSequence.Count;$k++){ $e=$s.TimeLine.MainSequence.Item($k); $effects += [ordered]@{index=$k; effect=$e.EffectType; trigger=$e.Timing.TriggerType; shape=$e.Shape.Name} } } catch {}
    $items += [ordered]@{slide=$i; shape_count=$s.Shapes.Count; animation_count=$s.TimeLine.MainSequence.Count; transition=$s.SlideShowTransition.EntryEffect; shapes=$shapes; effects=$effects}
  }
  $p.Close()
} finally { $app.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null }
$items | ConvertTo-Json -Depth 7 | Set-Content -Path $out -Encoding UTF8
Write-Output $out
