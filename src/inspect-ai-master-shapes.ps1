$input='C:\ppt-creater\production-template-library\ai-industry\master.pptx'
$out='C:\ppt-creater\catalog\ai-master-shape-inventory.json'
$app=New-Object -ComObject PowerPoint.Application;$app.Visible=-1
try {
  $p=$app.Presentations.Open($input,$true,$false,$false);$items=@()
  for($s=1;$s -le $p.Slides.Count;$s++){
    $shapes=@()
    for($i=1;$i -le $p.Slides.Item($s).Shapes.Count;$i++){
      $sh=$p.Slides.Item($s).Shapes.Item($i);$hasTf=$false;$text=''
      try{$hasTf=[bool]$sh.HasTextFrame;if($hasTf -and $sh.TextFrame.HasText){$text=$sh.TextFrame.TextRange.Text}}catch{}
      $shapes += [ordered]@{i=$i;name=$sh.Name;type=$sh.Type;has_text_frame=$hasTf;text=$text;left=[math]::Round($sh.Left);top=[math]::Round($sh.Top);width=[math]::Round($sh.Width);height=[math]::Round($sh.Height)}
    }
    $items += [ordered]@{slide=$s;shapes=$shapes}
  }
  $p.Close();$items|ConvertTo-Json -Depth 5|Set-Content $out -Encoding utf8
} finally {try{$app.Quit()}catch{};try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null}catch{}}
