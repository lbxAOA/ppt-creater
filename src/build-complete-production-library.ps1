# Build a strict-native production library from high-value source families.
# Each output master is a single-source PPTX subset; content validation replaces
# only existing nonempty text objects and preserves animations.
$ErrorActionPreference='Stop'
$repo='C:\ppt-creater';$sourceRoot=Join-Path $repo 'template-library\local-production-sources'
$productionRoot=Join-Path $repo 'complete-production-library';$validationRoot=Join-Path $productionRoot 'validation';New-Item -ItemType Directory -Force -Path $productionRoot,$validationRoot|Out-Null
$targets=[ordered]@{
 'ai-industry'=[ordered]@{source='modern-report-ai-investor-academic\source.pptx';slides=@(1,3,4,7,9,11,14,15,16,18,22,28)}
 'investor-pitch'=[ordered]@{source='modern-report-ai-investor-academic\source.pptx';slides=@(1,3,4,7,9,11,14,15,16,18,22,28)}
 'academic-clean'=[ordered]@{source='modern-report-ai-investor-academic\source.pptx';slides=@(1,3,4,7,9,11,14,15,16,18,22,28)}
 'data-boardroom'=[ordered]@{source='data-boardroom-corporate\source.pptx';slides=@(1,3,8,11,14,16,19,21,23,25,27)}
 'brand-company-profile'=[ordered]@{source='brand-company-profile\source.pptx';slides=@(1,2,3,5,7,9,11,13,15,17,19,21)}
 'marketing-event'=[ordered]@{source='marketing-event-culture\source.pptx';slides=@(1,2,3,4,5,6,7,8,9,10)}
 'culture-tourism'=[ordered]@{source='marketing-event-culture\source.pptx';slides=@(1,2,3,4,5,6,7,8,9,10)}
 'generic-corporate'=[ordered]@{source='data-boardroom-corporate\source.pptx';slides=@(1,3,8,11,14,16,19,21,23,25,27)}
 'wedding-bridal'=[ordered]@{source='wedding-bridal\source.pptx';slides=@(1,2,3,4,5,6,7,8,9,10)}
 'career-portfolio'=[ordered]@{source='career-portfolio\source.pptx';slides=@(1,3,5,7,9,11,13,15,17,19,21,23)}
}
function SafeText($old,$n){
 $limit=[Math]::Max(4,[Math]::Min(12,$old.Length))
 $samples=@('行业标题','核心判断','关键洞察','数据说明','行动建议','场景价值','增长路径','风险提示','客户价值','阶段成果','团队能力','结论')
 return $samples[($n-1)%$samples.Count].Substring(0,[Math]::Min($limit,$samples[($n-1)%$samples.Count].Length))
}
function Quit-App($app){try{$app.Quit()}catch{};try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null}catch{};[GC]::Collect();[GC]::WaitForPendingFinalizers()}
$result=[ordered]@{schema_version=1;overall_status='production_ready';targets=[ordered]@{}}
foreach($prop in $targets.GetEnumerator()){
 $name=$prop.Key;$cfg=$prop.Value;$src=Join-Path $sourceRoot $cfg.source;if(-not(Test-Path $src)){throw "Missing $src"}
 $app=$null
 try{
  $app=New-Object -ComObject PowerPoint.Application;$app.Visible=-1
  $dir=Join-Path $productionRoot $name;New-Item -ItemType Directory -Force -Path $dir|Out-Null;$master=Join-Path $dir 'master.pptx';Copy-Item $src $master -Force
  $p=$app.Presentations.Open($master,$false,$false,$false);$before=0;foreach($i in $cfg.slides){$before+=$p.Slides.Item($i).TimeLine.MainSequence.Count};for($i=$p.Slides.Count;$i -ge 1;$i--){if($cfg.slides -notcontains $i){$p.Slides.Item($i).Delete()}};$p.Save();try{$p.Close()}catch{}
  $m=$app.Presentations.Open($master,$false,$false,$false);for($ms=1;$ms -le $m.Slides.Count;$ms++){for($mi=1;$mi -le $m.Slides.Item($ms).Shapes.Count;$mi++){$msh=$m.Slides.Item($ms).Shapes.Item($mi);$mhas=$false;$mold='';try{$mhas=$msh.HasTextFrame -and $msh.TextFrame.HasText;if($mhas){$mold=$msh.TextFrame.TextRange.Text}}catch{};if($mhas -and $mold.Trim().Length -gt 1){$msh.Name=('slot_s{0:D2}_{1:D2}' -f $ms,$mi)}}};$m.Save();try{$m.Close()}catch{}
  $vdir=Join-Path $validationRoot $name;New-Item -ItemType Directory -Force -Path $vdir|Out-Null;$validation=Join-Path $vdir 'filled-validation.pptx';Copy-Item $master $validation -Force
  $v=$app.Presentations.Open($validation,$false,$false,$false);$fill=0;$slots=@();$pages=@()
  for($s=1;$s -le $v.Slides.Count;$s++){$pageSlots=@();for($i=1;$i -le $v.Slides.Item($s).Shapes.Count;$i++){$sh=$v.Slides.Item($s).Shapes.Item($i);$old='';$valid=$false;try{$valid=$sh.HasTextFrame -and $sh.TextFrame.HasText;if($valid){$old=$sh.TextFrame.TextRange.Text}}catch{};if($valid -and $old.Trim().Length -gt 0 -and $old.Length -gt 1){$slotName=('slot_s{0:D2}_{1:D2}' -f $s,$i);$sh.Name=$slotName;$sh.TextFrame.TextRange.Text=(SafeText $old ($fill+1));$slot=[ordered]@{name=$slotName;shape_index=$i;max_chars_cn=[Math]::Max(4,$old.Length);animation_locked=($v.Slides.Item($s).TimeLine.MainSequence.Count -gt 0)};$slots+=$slot;$pageSlots+=$slot;$fill++}};$pages += [ordered]@{page_index=$s;slots=$pageSlots;animation_effects=$v.Slides.Item($s).TimeLine.MainSequence.Count}}
  $v.Save();try{$v.Close()}catch{}
  $c=$app.Presentations.Open($validation,$true,$false,$false);$after=0;for($s=1;$s -le $c.Slides.Count;$s++){$after+=$c.Slides.Item($s).TimeLine.MainSequence.Count};$play=$true;$show=$null;try{$c.SlideShowSettings.ShowType=3;$show=$c.SlideShowSettings.Run();if($null -eq $show -or $null -eq $show.View){throw 'PowerPoint did not start a slide-show window'};Start-Sleep -Milliseconds 300;for($s=1;$s -le $c.Slides.Count;$s++){if($c.Slides.Item($s).TimeLine.MainSequence.Count -gt 0){$show.View.GotoSlide($s,$true);[void]$show.View.Next();Start-Sleep -Milliseconds 60}}}catch{$play=$false}finally{if($null -ne $show){try{$show.View.Exit()}catch{}}};try{$c.Close()}catch{}
  $result.targets[$name]=[ordered]@{master_pptx=$master;validation_pptx=$validation;source_pptx=$src;source_slide_indices=$cfg.slides;slide_count=$cfg.slides.Count;semantic_slots=$slots;pages=$pages;fill_count=$fill;new_shapes=0;strict_native_only=$true;animation_effects_before=$before;animation_effects_after=$after;powerpoint_playback=$play;status=if($before -eq $after -and $play -and $fill -gt 0){'production_ready'}else{'review_required'}}
 }finally{Quit-App $app}
}
if(@($result.targets.Values|Where-Object {$_.status -ne 'production_ready'}).Count -gt 0){$result.overall_status='review_required'}
$result|ConvertTo-Json -Depth 12|Set-Content (Join-Path $repo 'catalog\complete-production-template-library.json') -Encoding utf8
Write-Output (Join-Path $repo 'catalog\complete-production-template-library.json')
