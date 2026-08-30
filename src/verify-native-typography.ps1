$input='C:\ppt-creater\output\embodied-world-model-investor-briefing-native-typography.pptx'
$app=New-Object -ComObject PowerPoint.Application
$app.Visible=-1
try {
  $p=$app.Presentations.Open($input,$true,$false,$false)
  $r=@()
  for($s=1;$s -le $p.Slides.Count;$s++) {
    for($i=1;$i -le $p.Slides.Item($s).Shapes.Count;$i++) {
      $sh=$p.Slides.Item($s).Shapes.Item($i)
      if($sh.Name.StartsWith('injected_')) {
        $f=$sh.TextFrame.TextRange.Font
        $r += [ordered]@{slide=$s;name=$sh.Name;font=$f.Name;fontFE=$f.NameFarEast;size=$f.Size;color=$f.Color.RGB;bold=$f.Bold;alignment=$sh.TextFrame.TextRange.ParagraphFormat.Alignment}
      }
    }
  }
  $p.Close()
  $r | ConvertTo-Json -Depth 4 | Set-Content 'C:\ppt-creater\catalog\embodied-world-model-typography-check.json' -Encoding utf8
} finally {
  try{$app.Quit()}catch{}
  try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null}catch{}
}
