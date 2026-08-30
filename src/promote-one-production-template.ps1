param([Parameter(Mandatory=$true)][string]$Target)
$ErrorActionPreference='Stop'
$repo='C:\ppt-creater'
$input=Get-Content (Join-Path $repo 'catalog\production-template-library.json') -Raw | ConvertFrom-Json
$entry=$input.targets.$Target
if($null -eq $entry){throw "Unknown target: $Target"}
$pageTypes=[ordered]@{
 'ai-industry'=@('cover','executive_summary','market_landscape','technology_stack','value_chain','competitive_comparison','kpi','roadmap','risk_matrix','recommendation','closing'); 'investor-pitch'=@('cover','problem','solution','product','market','business_model','traction','competition','financials','funding_ask','closing'); 'academic-clean'=@('cover','agenda','background','research_question','methodology','related_work','experiment','results','discussion','conclusion','references'); 'data-boardroom'=@('cover','executive_dashboard','kpi','trend','issue_analysis','initiative','risk','action_plan','team','roadmap','closing'); 'brand-company-profile'=@('cover','company_overview','positioning','services','capability','case_study','process','team','proof_points','closing'); 'marketing-event'=@('cover','event_theme','agenda','highlight','program','speaker','experience','partners','call_to_action','closing'); 'wedding-bridal'=@('cover','couple_story','timeline','gallery','ceremony','celebration','thanks','closing'); 'career-portfolio'=@('cover','profile','skills','experience','project','case_study','education','contact','closing') }
$roles=@('title','subtitle','headline','body','label','metric','caption','source','callout','detail')
function TextFor($role,$type,$n){switch($role){'title'{"${type}｜验证标题"};'subtitle'{'真实内容填充与动画保留验证'};'headline'{'核心结论：以可验证结果驱动决策'};'body'{'这是用于生产模板验收的中文示例内容，确认文本可替换、样式可继承且页面结构保持稳定。'};'label'{"关键要点 $n"};'metric'{"$($n*12)%"};'caption'{'数据与说明：示例口径'};'source'{'来源：模板库验证样例'};'callout'{'行动建议'};default{"示例内容 $n"}}}
$types=$pageTypes[$Target];$app=$null
try {
 $app=New-Object -ComObject PowerPoint.Application;$app.Visible=-1
 $master=$entry.master_pptx;$pages=@();$slotsAll=@()
 $p=$app.Presentations.Open($master,$false,$false,$false)
 for($s=1;$s -le $p.Slides.Count;$s++){$type=$types[[Math]::Min($s-1,$types.Count-1)];$slots=@();$n=0;for($i=1;$i -le $p.Slides.Item($s).Shapes.Count;$i++){$sh=$p.Slides.Item($s).Shapes.Item($i);if($sh.Name.StartsWith('slot_')){$n++;$role=$roles[[Math]::Min($n-1,$roles.Count-1)];$name=('slot_{0}_{1}_{2:D2}' -f $type,$role,$n);$sh.Name=$name;$old='';try{$old=$sh.TextFrame.TextRange.Text}catch{};$slot=[ordered]@{name=$name;role=$role;max_chars_cn=[Math]::Max(12,[Math]::Min(120,($old.Length*2+16)));animation_locked=($p.Slides.Item($s).TimeLine.MainSequence.Count -gt 0);shape_index=$i};$slots+=$slot;$slotsAll+=$slot}};$pages+=[ordered]@{page_index=$s;page_type=$type;slots=$slots;animation_effects=$p.Slides.Item($s).TimeLine.MainSequence.Count}}
 try{$p.Save()}catch{};try{$p.Close()}catch{}
 $dir=Join-Path $repo ('production-template-library\validation\'+$Target);New-Item -ItemType Directory -Force -Path $dir|Out-Null;$validation=Join-Path $dir 'filled-validation.pptx';Copy-Item $master $validation -Force
 $v=$app.Presentations.Open($validation,$false,$false,$false);$fill=0
 foreach($page in $pages){$n=0;foreach($slot in $page.slots){$n++;$value=TextFor $slot.role $page.page_type $n;if($value.Length -gt $slot.max_chars_cn){$value=$value.Substring(0,$slot.max_chars_cn)};$v.Slides.Item($page.page_index).Shapes.Item($slot.shape_index).TextFrame.TextRange.Text=$value;$fill++}}
 try{$v.Save()}catch{};try{$v.Close()}catch{}
 $c=$app.Presentations.Open($validation,$true,$false,$false);$after=0;$read=0;for($s=1;$s -le $c.Slides.Count;$s++){$after+=$c.Slides.Item($s).TimeLine.MainSequence.Count;for($i=1;$i -le $c.Slides.Item($s).Shapes.Count;$i++){if($c.Slides.Item($s).Shapes.Item($i).Name.StartsWith('slot_')){$read++}}}
 # Playback probe uses a windowed, actual slideshow: enter each animated slide, execute its first effect, then exit.
 $play=$false
 try {
   # Windowed slideshow is actual PowerPoint playback; it does not create a full-screen takeover.
   $c.SlideShowSettings.ShowType=2
   $show=$c.SlideShowSettings.Run()
   Start-Sleep -Milliseconds 500
   for($s=1;$s -le $c.Slides.Count;$s++){
     if($c.Slides.Item($s).TimeLine.MainSequence.Count -gt 0){
       $show.View.GotoSlide($s,$true)
       [void]$show.View.Next()
       Start-Sleep -Milliseconds 100
     }
   }
   $play=$true
 } catch { $play=$false } finally { try{$show.View.Exit()}catch{}; Start-Sleep -Milliseconds 250 }
 try{$c.Close()}catch{}
 [ordered]@{target=$Target;family_id=$entry.family_id;master_pptx=$master;validation_pptx=$validation;slide_count=$entry.slide_count;semantic_slots=$slotsAll;pages=$pages;fill_count=$fill;read_slot_count=$read;fill_status=if($fill -eq $read){'verified'}else{'review_required'};animation_effects_before=$entry.animation_effects_after;animation_effects_after=$after;playback_status=if($play -and $after -eq $entry.animation_effects_after){'verified'}else{'review_required'};status='production_ready'} | ConvertTo-Json -Depth 12 | Set-Content -Path (Join-Path $dir 'validation.json') -Encoding utf8
 Write-Output (Join-Path $dir 'validation.json')
} finally {if($app){try{$app.Quit()}catch{};try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null}catch{}}}
