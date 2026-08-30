$ErrorActionPreference = 'Stop'
$repo = 'C:\ppt-creater'
$catalogPath = Join-Path $repo 'catalog\generation-template-library.json'
$productionPath = Join-Path $repo 'catalog\production-template-library.json'
$root = 'C:\PPT模板'
$targetRoot = Join-Path $repo 'production-template-library'
New-Item -ItemType Directory -Force -Path $targetRoot | Out-Null

$routes = Get-Content $catalogPath -Raw | ConvertFrom-Json
# Each page is copied inside its original family. Never use cross-template composition.
$selection = [ordered]@{
  'ai-industry' = [ordered]@{ family_id='02-005'; master='02_视觉风格\06_文艺唯美\模板\004_文艺唯美模板.pptx'; slides=@(1,2,3,4,7,10,13,16,19,22,25); note='Restrained editorial visual system; production candidate for AI/technical/strategy.' }
  'investor-pitch' = [ordered]@{ family_id='02-005'; master='02_视觉风格\06_文艺唯美\模板\004_文艺唯美模板.pptx'; slides=@(1,2,3,5,8,11,14,17,20,23,25); note='Same coherent family as AI industry; selected pages support narrative, evidence and closing.' }
  'academic-clean' = [ordered]@{ family_id='02-005'; master='02_视觉风格\06_文艺唯美\模板\004_文艺唯美模板.pptx'; slides=@(1,3,4,6,9,12,15,18,21,24,25); note='Clean research/teaching route; reviewer must check content density before filling.' }
  'data-boardroom' = [ordered]@{ family_id='04-008'; master='04_设计素材\04_汇报总结模板\模板\001_汇报总结模板.pptx'; slides=@(1,3,8,11,14,16,19,21,23,25,27); note='Native KPI, issue and action-review patterns for internal reporting.' }
  'brand-company-profile' = [ordered]@{ family_id='02-001'; master='02_视觉风格\01_动态风格\模板\065_动态模板.pptx'; slides=@(1,3,5,8,11,14,17,20,23,25); note='Editorial company-profile / proposal route.' }
  'marketing-event' = [ordered]@{ family_id='02-001'; master='02_视觉风格\01_动态风格\模板\065_动态模板.pptx'; slides=@(1,3,5,8,11,14,17,20,23,25); note='Editorial brand, launch and professional event route. Chinese-style cultural templates remain visual-reference assets until their editable text slots are mapped.' }
  'wedding-bridal' = [ordered]@{ family_id='03-007'; master='03_主题场景\02_婚礼相册\模板\003_婚礼相册模板.pptx'; slides=@(1,2,3,4,6,8,10,12); note='Photo-led vintage wedding / anniversary route.' }
  'career-portfolio' = [ordered]@{ family_id='01-006'; master='01_应用场景\04_求职与竞聘\模板\039_模板组.pptx'; slides=@(1,3,6,9,11,14,16,19,23); note='Personal resume, candidacy and portfolio route.' }
}

function Ensure-FamilySlots($presentation, $selectedSlides, $familyId) {
  foreach($idx in $selectedSlides) {
    $slide = $presentation.Slides.Item($idx)
    $slot=0
    for($i=1; $i -le $slide.Shapes.Count; $i++) {
      $shape=$slide.Shapes.Item($i)
      $hasText=$false; $text=''
      try { $hasText=$shape.HasTextFrame -and $shape.TextFrame.HasText; if($hasText){$text=$shape.TextFrame.TextRange.Text} } catch {}
      if($hasText) {
        $slot++
        # Reusable text targets retain their original object identity; only their name is normalized.
        if(-not $shape.Name.StartsWith('slot_')) { $shape.Name = ('slot_s{0:D2}_{1:D2}' -f $idx,$slot) }
      }
    }
  }
}

$app=New-Object -ComObject PowerPoint.Application
$app.Visible=-1
$output=[ordered]@{schema_version=1; generated_at=(Get-Date).ToString('s'); targets=[ordered]@{}}
try {
  foreach($target in $selection.Keys) {
    $entry=$selection[$target]
    $source=Join-Path $root $entry.master
    if(-not(Test-Path $source)){throw "Missing master for $target : $source"}
    $dir=Join-Path $targetRoot $target
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $out=Join-Path $dir 'master.pptx'
    Copy-Item $source $out -Force
    $p=$app.Presentations.Open($out,$false,$false,$false)
    $originalAnimations=@{}
    foreach($index in $entry.slides){ $originalAnimations[$index]=$p.Slides.Item($index).TimeLine.MainSequence.Count }
    # Delete non-selected slides bottom-up; preserved pages and their MainSequence remain untouched.
    for($i=$p.Slides.Count;$i -ge 1;$i--){if($entry.slides -notcontains $i){$p.Slides.Item($i).Delete()}}
    # Selected original slide references now appear in list order at 1..N.
    $renumbered=1..$entry.slides.Count
    Ensure-FamilySlots $p $renumbered $entry.family_id
    $animationAfter=0
    for($i=1;$i -le $p.Slides.Count;$i++){ $animationAfter += $p.Slides.Item($i).TimeLine.MainSequence.Count }
    $animationBefore=($originalAnimations.Values|Measure-Object -Sum).Sum
    $p.Save();$p.Close()
    $slotCount = 0
    $verify=$app.Presentations.Open($out,$true,$false,$false)
    for($i=1;$i -le $verify.Slides.Count;$i++){for($j=1;$j -le $verify.Slides.Item($i).Shapes.Count;$j++){if($verify.Slides.Item($i).Shapes.Item($j).Name.StartsWith('slot_')){$slotCount++}}}
    $verify.Close()
    $output.targets[$target]=[ordered]@{
      family_id=$entry.family_id; master_pptx=$out; source_pptx=$source; source_slide_indices=$entry.slides;
      slide_count=$entry.slides.Count; slot_count=$slotCount; animation_effects_before=$animationBefore; animation_effects_after=$animationAfter;
      animation_status=if($animationBefore -eq $animationAfter){'verified_preserved'}else{'review_required'}; status='template_ready_for_content_mapping'; note=$entry.note
    }
  }
} finally {$app.Quit();[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null;[GC]::Collect();[GC]::WaitForPendingFinalizers()}
$output | ConvertTo-Json -Depth 7 | Set-Content -Path $productionPath -Encoding utf8
Write-Output $productionPath
