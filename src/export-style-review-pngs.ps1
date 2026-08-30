$ErrorActionPreference='Stop'
$outDir='C:\ppt-creater\output\style-review-images';New-Item -ItemType Directory -Force -Path $outDir|Out-Null
$files=@(
 @{id='01-modern-report';file='C:\PPT模板\01_应用场景\02_汇报与答辩\模板\101_16比9比例.pptx'},
 @{id='02-data-boardroom';file='C:\PPT模板\04_设计素材\04_汇报总结模板\模板\001_汇报总结模板.pptx'},
 @{id='03-brand-profile';file='C:\PPT模板\02_视觉风格\07_扁平化\模板\015_扁平化模板.pptx'},
 @{id='04-event-culture';file='C:\PPT模板\03_主题场景\01_主题活动\模板\011_PPT模板宣传类.pptx'},
 @{id='05-wedding';file='C:\PPT模板\03_主题场景\01_主题活动\模板\003_PPT模板.pptx'},
 @{id='06-portfolio';file='C:\PPT模板\01_应用场景\04_求职与竞聘\模板\039_模板组.pptx'}
)
$app=New-Object -ComObject PowerPoint.Application;$app.Visible=-1
try{foreach($spec in $files){$p=$app.Presentations.Open($spec.file,$true,$false,$false);$dir=Join-Path $outDir $spec.id;New-Item -ItemType Directory -Force -Path $dir|Out-Null;for($s=1;$s -le $p.Slides.Count;$s++){$png=Join-Path $dir ('slide-{0:D3}.png' -f $s);$p.Slides.Item($s).Export($png,'PNG',1600,900)};$p.Close();Write-Output "$($spec.id) $($p.Slides.Count)"}}finally{try{$app.Quit()}catch{};try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null}catch{}}
