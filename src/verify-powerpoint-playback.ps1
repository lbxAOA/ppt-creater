param([Parameter(Mandatory=$true)][string]$FilePath,[Parameter(Mandatory=$true)][int]$Expected)
$app=New-Object -ComObject PowerPoint.Application;$app.Visible=-1
try {
 $p=$app.Presentations.Open($FilePath,$true,$false,$false);$count=0;for($s=1;$s -le $p.Slides.Count;$s++){$count+=$p.Slides.Item($s).TimeLine.MainSequence.Count}
 $ok=$true;try{$p.SlideShowSettings.ShowType=2;$show=$p.SlideShowSettings.Run();Start-Sleep -Milliseconds 400;for($s=1;$s -le $p.Slides.Count;$s++){if($p.Slides.Item($s).TimeLine.MainSequence.Count -gt 0){$show.View.GotoSlide($s,$true);[void]$show.View.Next();Start-Sleep -Milliseconds 80}};try{$show.View.Exit()}catch{}}catch{$ok=$false}
 try{$p.Close()}catch{}
 [ordered]@{input=$FilePath;effects=$count;expected=$Expected;playback=$ok;verified=($ok -and $count -eq $Expected)} | ConvertTo-Json -Compress
} finally {try{$app.Quit()}catch{};try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null}catch{}}
