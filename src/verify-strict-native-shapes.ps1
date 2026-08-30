$source='C:\PPT模板\01_应用场景\02_汇报与答辩\模板\101_16比9比例.pptx'
$output='C:\ppt-creater\output\embodied-world-model-investor-briefing-strict-native.pptx'
$selected=@(1,3,4,7,9,11,14,15,16,18,22,28)
$app=New-Object -ComObject PowerPoint.Application;$app.Visible=-1
try{
 $a=$app.Presentations.Open($source,$true,$false,$false);$b=$app.Presentations.Open($output,$true,$false,$false)
 $sourceCounts=@();$outCounts=@();foreach($idx in $selected){$sourceCounts+=$a.Slides.Item($idx).Shapes.Count};for($i=1;$i -le $b.Slides.Count;$i++){$outCounts+=$b.Slides.Item($i).Shapes.Count}
 $a.Close();$b.Close();[ordered]@{source_selected_shape_count=($sourceCounts|Measure-Object -Sum).Sum;output_shape_count=($outCounts|Measure-Object -Sum).Sum;shape_identity_preserved=(($sourceCounts|Measure-Object -Sum).Sum -eq ($outCounts|Measure-Object -Sum).Sum)}|ConvertTo-Json -Compress
}finally{try{$app.Quit()}catch{};try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null}catch{}}
