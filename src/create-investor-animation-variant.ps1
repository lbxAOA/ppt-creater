# Investor-animation variant: preserve all source slide objects; replace only animation sequences.
$ErrorActionPreference='Stop'
$source='C:\ppt-creater\output\embodied-world-model-investor-briefing-strict-native.pptx'
$output='C:\ppt-creater\output\embodied-world-model-investor-briefing-investor-animation.pptx'
$qa='C:\ppt-creater\catalog\embodied-world-model-investor-animation-qa.json'
Copy-Item $source $output -Force

# PowerPoint MsoAnimEffect constants: Fade=10. Build=1 (as one object), Trigger=1 (on click).
$EFFECT_FADE=10;$TRIGGER_ON_CLICK=1
$app=New-Object -ComObject PowerPoint.Application;$app.Visible=-1
try {
  $p=$app.Presentations.Open($output,$false,$false,$false)
  $sourceEffects=0;for($s=1;$s -le $p.Slides.Count;$s++){$sourceEffects += $p.Slides.Item($s).TimeLine.MainSequence.Count}
  # All target indices refer to existing original shapes. No shapes are added/deleted/repositioned.
  $plan=[ordered]@{
    4=@(1,2,3,4,5,6,7,8) # four investor theses, title + evidence pairs
    5=@(44,45,46,47,48,49,50,51,52,53) # perception → learning loop
    6=@(11,12,14,15,17,18,20,21) # four approaches: last is world-model loop
    7=@(13,14,15,16,17,18,19,20) # investor gates
    8=@(3,4,8,11,14) # four value statements
    9=@(5,6,7,8,9,10,11,12) # scorecard
    10=@(5,8,9,10,11,12) # deployment, data, ROI
    11=@(17,18,19,20,21,22,23,24) # risk cells
    12=@(1,6) # conclusion → closing
  }
  foreach($key in $plan.Keys) {
    $slide=$p.Slides.Item([int]$key)
    # Investor rhythm: remove inherited decorative effects, reveal only claim-bearing original objects.
    while($slide.TimeLine.MainSequence.Count -gt 0){$slide.TimeLine.MainSequence.Item(1).Delete()}
    foreach($shapeIndex in $plan[$key]) {
      $shape=$slide.Shapes.Item([int]$shapeIndex)
      [void]$slide.TimeLine.MainSequence.AddEffect($shape,$EFFECT_FADE,$TRIGGER_ON_CLICK)
    }
  }
  # Clear animation from non-content section pages to make transitions calmer.
  foreach($key in @(1,2,3)){ $slide=$p.Slides.Item($key); while($slide.TimeLine.MainSequence.Count -gt 0){$slide.TimeLine.MainSequence.Item(1).Delete()} }
  $p.Save();$p.Close()

  $v=$app.Presentations.Open($output,$true,$false,$false)
  $after=0;$countPlan=@{}
  for($s=1;$s -le $v.Slides.Count;$s++){$n=$v.Slides.Item($s).TimeLine.MainSequence.Count;$after += $n;$countPlan[$s]=$n}
  $play=$true
  try {
    $v.SlideShowSettings.ShowType=2
    $show=$v.SlideShowSettings.Run();Start-Sleep -Milliseconds 500
    foreach($slideNo in $plan.Keys | Sort-Object) {
      $show.View.GotoSlide([int]$slideNo,$true)
      $effectCount=$v.Slides.Item([int]$slideNo).TimeLine.MainSequence.Count
      for($i=1;$i -le $effectCount;$i++){[void]$show.View.Next();Start-Sleep -Milliseconds 100}
    }
    try{$show.View.Exit()}catch{}
  } catch {$play=$false}
  try{$v.Close()}catch{}
  [ordered]@{source_pptx=$source;output_pptx=$output;new_shapes=0;animation_style='native_object_fade_sequence';animation_effects_before=$sourceEffects;animation_effects_after=$after;per_slide_effects=($countPlan.GetEnumerator() | ForEach-Object {[ordered]@{slide=$_.Key;effects=$_.Value}});powerpoint_playback=$play;status=if($play){'verified'}else{'review_required'}} | ConvertTo-Json -Depth 5 | Set-Content $qa -Encoding utf8
} finally {try{$app.Quit()}catch{};try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null}catch{}}
Write-Output $output
