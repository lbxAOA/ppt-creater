# Promote user-approved round-4 network templates as isolated strict-native families.
# They supplement category routing but never mix with another source family.
$ErrorActionPreference='Stop'
$repo='C:\ppt-creater'
$catalog=Join-Path $repo 'catalog\network-production-template-library.json'
$specs=@(
 @{id='network-quarterly-business-review';sourceId='NET-10';input='C:\ppt-creater\network-supplements\NET-10\quarterly-business-review.pptx';targets=@('data-boardroom','generic-corporate','investor-pitch')},
 @{id='network-financial-management-consulting';sourceId='NET-12';input='C:\ppt-creater\network-supplements\NET-12\financial-management-consulting.pptx';targets=@('data-boardroom','generic-corporate','investor-pitch')}
)
function DemoText($old,$n){
  $pool=@('经营复盘','关键结论','目标达成','收入结构','增长驱动','核心挑战','行动计划','市场机会','客户价值','财务健康','风险应对','下一步安排')
  $value=$pool[($n-1)%$pool.Count]
  return $value.Substring(0,[Math]::Min([Math]::Max(2,$old.Length),$value.Length))
}
function Close-PowerPoint($app){
  if($null -ne $app){try{$app.Quit()}catch{};try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null}catch{}}
}
$data=Get-Content $catalog -Raw|ConvertFrom-Json
$app=New-Object -ComObject PowerPoint.Application
$app.Visible=-1
try {
  foreach($spec in $specs){
    $masterDir=Join-Path $repo ('network-production-library\'+$spec.id)
    $validationDir=Join-Path $repo ('network-production-library\validation\'+$spec.id)
    New-Item -ItemType Directory -Force -Path $masterDir,$validationDir|Out-Null
    $master=Join-Path $masterDir 'master.pptx'
    $validation=Join-Path $validationDir 'filled-validation.pptx'
    Copy-Item $spec.input $master -Force
    Copy-Item $master $validation -Force

    $p=$app.Presentations.Open($validation,$false,$false,$false)
    $before=0
    for($s=1;$s -le $p.Slides.Count;$s++){$before+=$p.Slides.Item($s).TimeLine.MainSequence.Count}
    $slideCount=$p.Slides.Count
    $slots=@();$pages=@();$fill=0
    for($s=1;$s -le $p.Slides.Count;$s++){
      $page=@();$isAnimationLocked=($p.Slides.Item($s).TimeLine.MainSequence.Count -gt 0)
      for($i=1;$i -le $p.Slides.Item($s).Shapes.Count;$i++){
        $shape=$p.Slides.Item($s).Shapes.Item($i);$old='';$editable=$false
        try {$editable=$shape.HasTextFrame -and $shape.TextFrame.HasText;if($editable){$old=$shape.TextFrame.TextRange.Text}}catch{}
        if($editable -and $old.Trim().Length -gt 1){
          $name=('slot_s{0:D2}_{1:D2}' -f $s,$i)
          $shape.Name=$name
          $shape.TextFrame.TextRange.Text=(DemoText $old ($fill+1))
          $slot=[ordered]@{name=$name;shape_index=$i;max_chars_cn=$old.Length;animation_locked=$isAnimationLocked}
          $slots+=$slot;$page+=$slot;$fill++
        }
      }
      $pages+=[ordered]@{source_slide_index=$s;slots=$page;animation_effects=$p.Slides.Item($s).TimeLine.MainSequence.Count}
    }
    $p.Save();$p.Close()

    $v=$app.Presentations.Open($validation,$true,$false,$false)
    $after=0
    for($s=1;$s -le $v.Slides.Count;$s++){$after+=$v.Slides.Item($s).TimeLine.MainSequence.Count}
    $play=$true;$show=$null
    try {
      $v.SlideShowSettings.ShowType=2
      $show=$v.SlideShowSettings.Run()
      if($null -eq $show -or $null -eq $show.View){throw 'PowerPoint did not start a slide-show window'}
      Start-Sleep -Milliseconds 300
      $show.View.GotoSlide(1,$true)
      Start-Sleep -Milliseconds 100
    } catch {$play=$false
    } finally {if($null -ne $show){try{$show.View.Exit()}catch{}}}
    try{$v.Close()}catch{}

    $item=[ordered]@{
      source_id=$spec.sourceId;source_pptx=$spec.input;master_pptx=$master;validation_pptx=$validation
      route_targets=$spec.targets;slide_count=$slideCount;semantic_slots=$slots;pages=$pages;fill_count=$fill
      new_shapes=0;strict_native_only=$true;animation_effects_before=$before;animation_effects_after=$after
      powerpoint_playback=$play;license_status='manual_external_review_required'
      status=if($play -and $before -eq $after -and $fill -gt 0){'production_ready'}else{'review_required'}
    }
    $data.families|Add-Member -Force -NotePropertyName $spec.id -NotePropertyValue $item
  }
  $data.overall_status=if(@($data.families.psobject.Properties|Where-Object {$_.Value.status -ne 'production_ready'}).Count -eq 0){'production_ready'}else{'review_required'}
  $data|ConvertTo-Json -Depth 12|Set-Content $catalog -Encoding utf8
} finally {Close-PowerPoint $app}
Write-Output $catalog
