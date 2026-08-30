$ErrorActionPreference='Stop'
$files=@(
 @{id='NET-02';file='C:\ppt-creater\network-supplements\NET-02\simple-professional.pptx'},
 @{id='NET-03';file='C:\ppt-creater\network-supplements\NET-03\my-portfolio.pptx'},
 @{id='NET-04';file='C:\ppt-creater\network-supplements\NET-04\amelia-wedding.pptx'}
)
$app=New-Object -ComObject PowerPoint.Application;$app.Visible=-1
try{foreach($spec in $files){$p=$app.Presentations.Open($spec.file,$true,$false,$false);$dir=Join-Path 'C:\ppt-creater\network-supplements' (Join-Path $spec.id 'previews');New-Item -ItemType Directory -Force -Path $dir|Out-Null;for($s=1;$s -le $p.Slides.Count;$s++){$p.Slides.Item($s).Export((Join-Path $dir ('slide-{0:D3}.png' -f $s)),'PNG',1600,900)};$n=$p.Slides.Count;$p.Close();Write-Output "$($spec.id) $n"}}finally{try{$app.Quit()}catch{};try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null}catch{}}
