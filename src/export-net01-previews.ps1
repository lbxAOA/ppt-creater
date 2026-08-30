$input='C:\ppt-creater\network-supplements\NET-01\pitch-with-confidence.pptx'
$out='C:\ppt-creater\network-supplements\NET-01\previews';New-Item -ItemType Directory -Force -Path $out|Out-Null
$app=New-Object -ComObject PowerPoint.Application;$app.Visible=-1
try{$p=$app.Presentations.Open($input,$true,$false,$false);for($s=1;$s -le $p.Slides.Count;$s++){$p.Slides.Item($s).Export((Join-Path $out ('slide-{0:D3}.png' -f $s)),'PNG',1600,900)};$count=$p.Slides.Count;$p.Close();Write-Output $count}finally{try{$app.Quit()}catch{};try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null}catch{}}
