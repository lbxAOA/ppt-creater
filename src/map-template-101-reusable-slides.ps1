$input='C:\PPT模板\01_应用场景\02_汇报与答辩\模板\101_16比9比例.pptx'
$out='C:\ppt-creater\catalog\template-101-reusable-slide-map.json'
$app=New-Object -ComObject PowerPoint.Application;$app.Visible=-1
try{
 $p=$app.Presentations.Open($input,$true,$false,$false)
 $selected=@(1,3,4,7,9,11,13,14,15,16,18,22,27,28)
 $items=@()
 foreach($s in $selected){
  $slots=@();$slide=$p.Slides.Item($s)
  for($i=1;$i -le $slide.Shapes.Count;$i++){
   $sh=$slide.Shapes.Item($i);$text='';$has=$false
   try{$has=[bool]$sh.HasTextFrame;if($has -and $sh.TextFrame.HasText){$text=$sh.TextFrame.TextRange.Text}}catch{}
   if($has -and $text.Trim().Length -gt 0 -and $text -notmatch '^(P|T|\+|\d+|ART\s+\d+)$'){
     $f=$sh.TextFrame.TextRange.Font
     $slots += [ordered]@{shape_index=$i;shape_name=$sh.Name;source_text=$text;source_chars=$text.Length;max_chars_cn=[Math]::Max(8,[Math]::Min(140,($text.Length*2)));font=$f.Name;font_fe=$f.NameFarEast;font_size=$f.Size;color=$f.Color.RGB;bold=$f.Bold;animation_locked=($slide.TimeLine.MainSequence.Count -gt 0)}
   }
  }
  $items += [ordered]@{source_slide=$s;animation_effects=$slide.TimeLine.MainSequence.Count;slots=$slots}
 }
 $p.Close();$items|ConvertTo-Json -Depth 6|Set-Content $out -Encoding utf8
}finally{try{$app.Quit()}catch{};try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null}catch{}}
