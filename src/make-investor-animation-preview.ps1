# Animation preview: retain the native template's existing animation targets and
# normalize timing into deliberate investor-report reveal groups. No shapes added,
# deleted, resized or reformatted.
$ErrorActionPreference='Stop'
$source='C:\ppt-creater\output\embodied-world-model-investor-briefing-strict-native.pptx'
$output='C:\ppt-creater\output\embodied-world-model-investor-briefing-investor-animation-preview.pptx'
$qa='C:\ppt-creater\catalog\embodied-world-model-investor-animation-qa.json'
Copy-Item $source $output -Force
$PPT_TRIGGER_ON_CLICK=1;$PPT_TRIGGER_WITH_PREVIOUS=2;$PPT_TRIGGER_AFTER_PREVIOUS=3
$app=New-Object -ComObject PowerPoint.Application;$app.Visible=-1
try{
 $p=$app.Presentations.Open($output,$false,$false,$false)
 $before=0;for($s=1;$s -le $p.Slides.Count;$s++){$before+=$p.Slides.Item($s).TimeLine.MainSequence.Count}
 # Slide 4: four investor thesis cards — title then body, one click per card.
 $slide=$p.Slides.Item(4);$seq=$slide.TimeLine.MainSequence
 for($i=1;$i -le $seq.Count;$i++){$e=$seq.Item($i);$e.Timing.Duration=0.35;$e.Timing.TriggerDelayTime=0;if($i % 2 -eq 1){$e.Timing.TriggerType=$PPT_TRIGGER_ON_CLICK}else{$e.Timing.TriggerType=$PPT_TRIGGER_WITH_PREVIOUS}}
 # Slide 7: four diligence gates — each gate's visual and copy become one click group.
 $slide=$p.Slides.Item(7);$seq=$slide.TimeLine.MainSequence
 for($i=1;$i -le $seq.Count;$i++){$e=$seq.Item($i);$e.Timing.Duration=0.3;$e.Timing.TriggerDelayTime=0;if($i -in @(1,5,9,13)){$e.Timing.TriggerType=$PPT_TRIGGER_ON_CLICK}else{$e.Timing.TriggerType=$PPT_TRIGGER_WITH_PREVIOUS}}
 # Slide 8: four investment-value panels — each panel is a single reveal group.
 $slide=$p.Slides.Item(8);$seq=$slide.TimeLine.MainSequence
 for($i=1;$i -le $seq.Count;$i++){$e=$seq.Item($i);$e.Timing.Duration=0.3;$e.Timing.TriggerDelayTime=0;if($i -in @(1,4,8,12)){$e.Timing.TriggerType=$PPT_TRIGGER_ON_CLICK}else{$e.Timing.TriggerType=$PPT_TRIGGER_WITH_PREVIOUS}}
 # Slide 9: four scorecard labels and metrics reveal as concise decision gates.
 $slide=$p.Slides.Item(9);$seq=$slide.TimeLine.MainSequence
 for($i=1;$i -le $seq.Count;$i++){$e=$seq.Item($i);$e.Timing.Duration=0.35;$e.Timing.TriggerDelayTime=0;if($i -in @(1,3,5,7)){$e.Timing.TriggerType=$PPT_TRIGGER_ON_CLICK}else{$e.Timing.TriggerType=$PPT_TRIGGER_WITH_PREVIOUS}}
 $p.Save();$p.Close()
 $v=$app.Presentations.Open($output,$true,$false,$false);$after=0;for($s=1;$s -le $v.Slides.Count;$s++){$after+=$v.Slides.Item($s).TimeLine.MainSequence.Count}
 $play=$true
 $show=$null
 try{
   $v.SlideShowSettings.ShowType=2
   $show=$v.SlideShowSettings.Run()
   if($null -eq $show -or $null -eq $show.View){throw 'PowerPoint did not start a slide-show window'}
   Start-Sleep -Milliseconds 400
   for($s=1;$s -le $v.Slides.Count;$s++){
     $n=$v.Slides.Item($s).TimeLine.MainSequence.Count
     if($n -gt 0){
       $show.View.GotoSlide($s,$true)
       for($i=1;$i -le $n;$i++){[void]$show.View.Next();Start-Sleep -Milliseconds 70}
     }
   }
 }catch{
   $play=$false
 }finally{
   if($null -ne $show){try{$show.View.Exit()}catch{}}
 }
 try{$v.Close()}catch{}
 $qaData=[ordered]@{
  source_pptx=$source
  output_pptx=$output
  new_shapes=0
  animation_style='native_object_fade_sequence'
  animation_effects_before=$before
  animation_effects_after=$after
  reworked_slides=@(4,7,8,9)
  animation_policy='0.30-0.35s, one click per investment argument, supporting elements with previous'
  powerpoint_playback=$play
  status=if($before -eq $after -and $play){'verified'}else{'review_required'}
}
$qaData | ConvertTo-Json -Depth 5 | Set-Content $qa -Encoding utf8
}finally{try{$app.Quit()}catch{};try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null}catch{}}
Write-Output $output
