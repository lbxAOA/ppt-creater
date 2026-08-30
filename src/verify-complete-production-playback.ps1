# Re-check every completed validation deck with PowerPoint and update the registry
# only with observed playback results. This does not alter any PPTX.
$ErrorActionPreference='Stop'
$registryPath='C:\ppt-creater\catalog\complete-production-template-library.json'
$registry=Get-Content $registryPath -Raw|ConvertFrom-Json
$app=New-Object -ComObject PowerPoint.Application;$app.Visible=-1
try {
 foreach($property in $registry.targets.psobject.Properties){
   $item=$property.Value;$p=$null;$play=$false
   try{
     $p=$app.Presentations.Open($item.validation_pptx,$true,$false,$false);$count=0;for($s=1;$s -le $p.Slides.Count;$s++){$count+=$p.Slides.Item($s).TimeLine.MainSequence.Count}
     $p.SlideShowSettings.ShowType=2;$show=$p.SlideShowSettings.Run();Start-Sleep -Milliseconds 400
     for($s=1;$s -le $p.Slides.Count;$s++){if($p.Slides.Item($s).TimeLine.MainSequence.Count -gt 0){$show.View.GotoSlide($s,$true);[void]$show.View.Next();Start-Sleep -Milliseconds 75}}
     try{$show.View.Exit()}catch{};$play=($count -eq $item.animation_effects_before)
   }catch{Write-Output ("FAIL " + $property.Name + ": " + $_.Exception.Message);$play=$false}finally{try{if($p){$p.Close()}}catch{}}
   $item.powerpoint_playback=$play
   if($item.strict_native_only -and $item.new_shapes -eq 0 -and $item.animation_effects_before -eq $item.animation_effects_after -and $play -and $item.fill_count -gt 0){$item.status='production_ready'}else{$item.status='review_required'}
 }
 $blocked=@($registry.targets.psobject.Properties|Where-Object {$_.Value.status -ne 'production_ready'})
 $registry.overall_status=if($blocked.Count -eq 0){'production_ready'}else{'review_required'}
 $registry|ConvertTo-Json -Depth 12|Set-Content $registryPath -Encoding utf8
}finally{try{$app.Quit()}catch{};try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null}catch{}}
Get-Content $registryPath -Raw|ConvertFrom-Json|Select-Object -ExpandProperty overall_status
