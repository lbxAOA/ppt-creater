$templates=@(
 'C:\PPT模板\02_视觉风格\05_中国风\模板\030_中国风模板.pptx',
 'C:\PPT模板\02_视觉风格\05_中国风\模板\052_中国风模板.pptx'
)
$app=New-Object -ComObject PowerPoint.Application;$app.Visible=-1
try{foreach($input in $templates){$p=$app.Presentations.Open($input,$true,$false,$false);Write-Output "TEMPLATE $input";for($s=1;$s -le $p.Slides.Count;$s++){$items=@();for($i=1;$i -le $p.Slides.Item($s).Shapes.Count;$i++){$sh=$p.Slides.Item($s).Shapes.Item($i);$has=$false;$txt='';try{$has=$sh.HasTextFrame;if($has -and $sh.TextFrame.HasText){$txt=$sh.TextFrame.TextRange.Text}}catch{};if($has){$items += "${i}:$($txt.Substring(0,[Math]::Min(40,$txt.Length)))"}};if($items.Count){Write-Output "Slide $s :: $($items -join ' | ')"}};$p.Close()}}finally{try{$app.Quit()}catch{};try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null}catch{}}
