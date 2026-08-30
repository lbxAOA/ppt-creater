$input='C:\PPT模板\01_应用场景\02_汇报与答辩\模板\101_16比9比例.pptx'
$out='C:\ppt-creater\catalog\template-101-page-inventory.json'
$app=New-Object -ComObject PowerPoint.Application;$app.Visible=-1
try{
 $p=$app.Presentations.Open($input,$true,$false,$false);$items=@()
 for($s=1;$s -le $p.Slides.Count;$s++){
  $texts=@();$pictures=@();$shapeCount=$p.Slides.Item($s).Shapes.Count
  for($i=1;$i -le $shapeCount;$i++){
   $sh=$p.Slides.Item($s).Shapes.Item($i);$txt='';$has=$false
   try{$has=[bool]$sh.HasTextFrame;if($has -and $sh.TextFrame.HasText){$txt=$sh.TextFrame.TextRange.Text}}catch{}
   if($has -and $txt.Trim().Length -gt 0){$f=$sh.TextFrame.TextRange.Font;$texts += [ordered]@{i=$i;name=$sh.Name;text=$txt;chars=$txt.Length;left=[math]::Round($sh.Left);top=[math]::Round($sh.Top);width=[math]::Round($sh.Width);height=[math]::Round($sh.Height);font=$f.Name;font_fe=$f.NameFarEast;size=$f.Size;color=$f.Color.RGB;bold=$f.Bold}}
   if($sh.Type -eq 13){$pictures += [ordered]@{i=$i;name=$sh.Name;left=[math]::Round($sh.Left);top=[math]::Round($sh.Top);width=[math]::Round($sh.Width);height=[math]::Round($sh.Height)}}
  }
  $items += [ordered]@{slide=$s;shape_count=$shapeCount;effects=$p.Slides.Item($s).TimeLine.MainSequence.Count;texts=$texts;pictures=$pictures}
 }
 $p.Close();$items|ConvertTo-Json -Depth 6|Set-Content $out -Encoding utf8
}finally{try{$app.Quit()}catch{};try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null}catch{}}
