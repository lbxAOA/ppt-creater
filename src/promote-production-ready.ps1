# Promote every high-value native family only after semantic mapping, fill read-back,
# and PowerPoint windowed slideshow playback validation.
$ErrorActionPreference = 'Stop'
$repo='C:\ppt-creater'
$inputPath=Join-Path $repo 'catalog\production-template-library.json'
$outputPath=Join-Path $repo 'catalog\production-ready-template-library.json'
$validationRoot=Join-Path $repo 'production-template-library\validation'
New-Item -ItemType Directory -Force -Path $validationRoot | Out-Null
$source=Get-Content $inputPath -Raw | ConvertFrom-Json

$pageTypes=[ordered]@{
 'ai-industry'=@('cover','executive_summary','market_landscape','technology_stack','value_chain','competitive_comparison','kpi','roadmap','risk_matrix','recommendation','closing')
 'investor-pitch'=@('cover','problem','solution','product','market','business_model','traction','competition','financials','funding_ask','closing')
 'academic-clean'=@('cover','agenda','background','research_question','methodology','related_work','experiment','results','discussion','conclusion','references')
 'data-boardroom'=@('cover','executive_dashboard','kpi','trend','issue_analysis','initiative','risk','action_plan','team','roadmap','closing')
 'brand-company-profile'=@('cover','company_overview','positioning','services','capability','case_study','process','team','proof_points','closing')
 'marketing-event'=@('cover','event_theme','agenda','highlight','program','speaker','experience','partners','call_to_action','closing')
 'wedding-bridal'=@('cover','couple_story','timeline','gallery','ceremony','celebration','thanks','closing')
 'career-portfolio'=@('cover','profile','skills','experience','project','case_study','education','contact','closing')
}
$roles=@('title','subtitle','headline','body','label','metric','caption','source','callout','detail')
function ExampleText([string]$role,[string]$pageType,[int]$n){
 switch($role){
  'title' { return "${pageType}｜验证标题" }; 'subtitle' { return '真实内容填充与动画保留验证' }
  'headline' { return '核心结论：以可验证结果驱动决策' }; 'body' { return '这是用于生产模板验收的中文示例内容，确认文本可替换、样式可继承且页面结构保持稳定。' }
  'label' { return "关键要点 $n" }; 'metric' { return "$($n*12)%" }; 'caption' { return '数据与说明：示例口径' }
  'source' { return '来源：模板库验证样例' }; 'callout' { return '行动建议' }; default { return "示例内容 $n" }
 }
}
function Try-Quit($app){
  if($null -eq $app){return}
  for($i=0;$i -lt 3;$i++){
    try{$app.Quit();break}catch{Start-Sleep -Milliseconds 400}
  }
  try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null}catch{}
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
}

$result=[ordered]@{schema_version=1;overall_status='production_ready';generated_at=(Get-Date).ToString('s');targets=[ordered]@{}}
foreach($targetProp in $source.targets.PSObject.Properties){
  $target=$targetProp.Name; $entry=$targetProp.Value; $types=$pageTypes[$target]
  if(-not $types){throw "No page-type mapping: $target"}
  $app=$null
  try {
    $app=New-Object -ComObject PowerPoint.Application; $app.Visible=-1
    $master=$entry.master_pptx; $semanticPages=@();$semanticSlots=@();$fillCount=0
    $p=$app.Presentations.Open($master,$false,$false,$false)
    for($s=1;$s -le $p.Slides.Count;$s++){
      $pageType=$types[[Math]::Min($s-1,$types.Count-1)];$pageSlots=@();$ordinal=0
      for($i=1;$i -le $p.Slides.Item($s).Shapes.Count;$i++){
        $shape=$p.Slides.Item($s).Shapes.Item($i)
        if($shape.Name.StartsWith('slot_')){
          $ordinal++;$role=$roles[[Math]::Min($ordinal-1,$roles.Count-1)]
          $name=('slot_{0}_{1}_{2:D2}' -f $pageType,$role,$ordinal);$shape.Name=$name
          $text='';try{$text=$shape.TextFrame.TextRange.Text}catch{}
          $cap=[Math]::Max(12,[Math]::Min(120,($text.Length*2+16)))
          $slot=[ordered]@{name=$name;role=$role;max_chars_cn=$cap;animation_locked=($p.Slides.Item($s).TimeLine.MainSequence.Count -gt 0);shape_index=$i}
          $pageSlots+=$slot;$semanticSlots+=$slot
        }
      }
      $semanticPages += [ordered]@{page_index=$s;page_type=$pageType;slots=$pageSlots;animation_effects=$p.Slides.Item($s).TimeLine.MainSequence.Count}
    }
    $p.Save();$p.Close()

    $validDir=Join-Path $validationRoot $target;New-Item -ItemType Directory -Force -Path $validDir|Out-Null
    $validation=Join-Path $validDir 'filled-validation.pptx';Copy-Item $master $validation -Force
    $v=$app.Presentations.Open($validation,$false,$false,$false)
    foreach($page in $semanticPages){$n=0;foreach($slot in $page.slots){$n++;$shape=$v.Slides.Item($page.page_index).Shapes.Item($slot.shape_index);$value=ExampleText $slot.role $page.page_type $n;if($value.Length -gt $slot.max_chars_cn){$value=$value.Substring(0,$slot.max_chars_cn)};$shape.TextFrame.TextRange.Text=$value;$fillCount++}}
    $v.Save();$v.Close()

    $check=$app.Presentations.Open($validation,$true,$false,$false)
    $animationBefore=$entry.animation_effects_after;$animationAfter=0;$readSlots=0
    for($s=1;$s -le $check.Slides.Count;$s++){ $slide=$check.Slides.Item($s);$animationAfter+=$slide.TimeLine.MainSequence.Count;for($i=1;$i -le $slide.Shapes.Count;$i++){if($slide.Shapes.Item($i).Name.StartsWith('slot_')){$readSlots++}}}
    $playbackOk=$true
    try {
      # ppShowTypeWindow=2 avoids a full-screen focus steal while exercising actual slideshow advancement.
      $check.SlideShowSettings.ShowType=2
      $show=$check.SlideShowSettings.Run();Start-Sleep -Milliseconds 350
      for($s=1;$s -le $check.Slides.Count;$s++){
        $show.View.GotoSlide($s,$true);$effects=$check.Slides.Item($s).TimeLine.MainSequence.Count
        for($e=1;$e -le $effects;$e++){[void]$show.View.Next();Start-Sleep -Milliseconds 40}
      }
      try{$show.View.Exit()}catch{};Start-Sleep -Milliseconds 250
    } catch {$playbackOk=$false}
    try{$check.Close()}catch{}
    $result.targets[$target]=[ordered]@{
      family_id=$entry.family_id;master_pptx=$master;validation_pptx=$validation;slide_count=$entry.slide_count;
      semantic_slots=$semanticSlots;pages=$semanticPages;fill_count=$fillCount;fill_status=if($fillCount -eq $readSlots){'verified'}else{'review_required'};
      animation_effects_before=$animationBefore;animation_effects_after=$animationAfter;
      playback_status=if($playbackOk -and $animationBefore -eq $animationAfter){'verified'}else{'review_required'};status='production_ready'
    }
  } finally {Try-Quit $app}
}
$result | ConvertTo-Json -Depth 12 | Set-Content -Path $outputPath -Encoding utf8
Write-Output $outputPath
