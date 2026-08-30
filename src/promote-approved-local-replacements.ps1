# Promote the user-approved replacement and addition families from curated copies.
# All sources are project-contained before this script runs. Each family remains
# single-source; validation replaces only existing native text shapes.
$ErrorActionPreference='Stop'
$repo='C:\ppt-creater'
$sourceRoot=Join-Path $repo 'curated\template-families'
$productionRoot=Join-Path $repo 'approved-template-additions'
$validationRoot=Join-Path $productionRoot 'validation'
New-Item -ItemType Directory -Force -Path $productionRoot,$validationRoot|Out-Null

$families=[ordered]@{
 'modern-report-ai-investor-academic'=[ordered]@{source='02-005\004_文艺唯美模板.pptx';slides=@(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25);route_targets=@('ai-industry','investor-pitch','academic-clean')}
 'marketing-event-chinese-style'=[ordered]@{source='02-004\052_中国风模板.pptx';slides=@(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25);route_targets=@('marketing-event')}
 'culture-tourism-chinese-style'=[ordered]@{source='02-003\030_中国风模板.pptx';slides=@(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20);route_targets=@('culture-tourism')}
 'wedding-album-expanded'=[ordered]@{source='03-004\004_婚礼相册模板.pptx';slides=@(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43);route_targets=@('wedding-bridal')}
 'data-visual-reference'=[ordered]@{source='04-004\013_精品图表精品推荐.pptx';slides=@(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,96,97,98,99,100,101,102,103,104,105,106,107,108,109,110,111,112,113,114,115,116,117,118,119,120,121,122,123,124,125,126,127,128,129,130,131,132,133,134,135,136,137,138,139,140,141,142,143,144,145,146,147,148,149,150,151,152,153,154,155,156,157,158,159,160,161,162,163,164,165,166,167,168,169,170,171,172,173,174);route_targets=@('ai-industry','data-boardroom');usage='reference_only'}
 'data-chart-reference'=[ordered]@{source='04-006\007_常用图表.pptx';slides=@(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38);route_targets=@('ai-industry','data-boardroom');usage='reference_only'}
}

function SafeText($old,$n){
 $limit=[Math]::Max(4,[Math]::Min(12,$old.Length))
 $samples=@('行业标题','核心判断','关键洞察','数据说明','行动建议','场景价值','增长路径','风险提示','客户价值','阶段成果','团队能力','结论')
 return $samples[($n-1)%$samples.Count].Substring(0,[Math]::Min($limit,$samples[($n-1)%$samples.Count].Length))
}
function Quit-App($app){try{$app.Quit()}catch{};try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null}catch{};[GC]::Collect();[GC]::WaitForPendingFinalizers()}

$result=[ordered]@{schema_version=1;overall_status='production_ready';families=[ordered]@{}}
foreach($prop in $families.GetEnumerator()){
 $name=$prop.Key;$cfg=$prop.Value;$src=Join-Path $sourceRoot $cfg.source
 if(-not(Test-Path $src)){throw "Missing project-contained source $src"}
 $app=$null
 try{
  $app=New-Object -ComObject PowerPoint.Application;$app.Visible=-1
  $dir=Join-Path $productionRoot $name;New-Item -ItemType Directory -Force -Path $dir|Out-Null;$master=Join-Path $dir 'master.pptx';Copy-Item $src $master -Force
  $p=$app.Presentations.Open($master,$false,$false,$false);$before=0;foreach($i in $cfg.slides){$before += $p.Slides.Item($i).TimeLine.MainSequence.Count};for($i=$p.Slides.Count;$i -ge 1;$i--){if($cfg.slides -notcontains $i){$p.Slides.Item($i).Delete()}};$p.Save();try{$p.Close()}catch{}
  $vdir=Join-Path $validationRoot $name;New-Item -ItemType Directory -Force -Path $vdir|Out-Null;$validation=Join-Path $vdir 'filled-validation.pptx';Copy-Item $master $validation -Force
  $v=$app.Presentations.Open($validation,$false,$false,$false);$fill=0;$slots=@();$pages=@()
  for($s=1;$s -le $v.Slides.Count;$s++){$pageSlots=@();for($i=1;$i -le $v.Slides.Item($s).Shapes.Count;$i++){$sh=$v.Slides.Item($s).Shapes.Item($i);$old='';$hasTextFrame=$false;try{$hasTextFrame=$sh.HasTextFrame;if($hasTextFrame){$old=$sh.TextFrame.TextRange.Text}}catch{};$canFill=$hasTextFrame -and (($old.Trim().Length -gt 1) -or ($old.Trim().Length -eq 0 -and $fill -lt 8));if($canFill){$slotName=('slot_s{0:D3}_{1:D3}' -f $s,$i);$sh.Name=$slotName;$sh.TextFrame.TextRange.Text=(SafeText $old ($fill+1));$slot=[ordered]@{name=$slotName;shape_index=$i;max_chars_cn=[Math]::Max(4,[Math]::Max($old.Length,12));animation_locked=($v.Slides.Item($s).TimeLine.MainSequence.Count -gt 0)};$slots+=$slot;$pageSlots+=$slot;$fill++}};$pages += [ordered]@{page_index=$s;slots=$pageSlots;animation_effects=$v.Slides.Item($s).TimeLine.MainSequence.Count}}
  $v.Save();try{$v.Close()}catch{}
  # PowerPoint may rewrite animation XML after text replacement. Restore only the
  # raw timing/transition blocks from the native master; edited shapes stay intact.
  $timed=Join-Path $vdir 'filled-validation-timed.pptx'
  & python (Join-Path $repo 'src\ppt_creater\restore_animation_timing.py') $master $validation $timed
  if($LASTEXITCODE -ne 0){throw 'Animation timing restoration failed'}
  Move-Item $timed $validation -Force
  $c=$app.Presentations.Open($validation,$true,$false,$false);$after=0;for($s=1;$s -le $c.Slides.Count;$s++){$after+=$c.Slides.Item($s).TimeLine.MainSequence.Count};$play=$true;$show=$null;try{$c.SlideShowSettings.ShowType=3;$show=$c.SlideShowSettings.Run();if($null -eq $show -or $null -eq $show.View){throw 'PowerPoint did not start a slide-show window'};Start-Sleep -Milliseconds 300;for($s=1;$s -le $c.Slides.Count;$s++){if($c.Slides.Item($s).TimeLine.MainSequence.Count -gt 0){$show.View.GotoSlide($s,$true);[void]$show.View.Next();Start-Sleep -Milliseconds 40}}}catch{$play=$false}finally{if($null -ne $show){try{$show.View.Exit()}catch{}}};try{$c.Close()}catch{}
  $item=[ordered]@{source_pptx=$src;master_pptx=$master;validation_pptx=$validation;route_targets=$cfg.route_targets;usage=if($cfg.usage){$cfg.usage}else{'production_replacement'};slide_count=$cfg.slides.Count;semantic_slots=$slots;pages=$pages;fill_count=$fill;new_shapes=0;strict_native_only=$true;animation_effects_before=$before;animation_effects_after=$after;powerpoint_playback=$play;status=if($before -eq $after -and $play -and $fill -gt 0){'production_ready'}else{'review_required'}}
  $result.families[$name]=$item
 }finally{Quit-App $app}
}
if(@($result.families.Values|Where-Object {$_.status -ne 'production_ready'}).Count -gt 0){$result.overall_status='review_required'}
$result|ConvertTo-Json -Depth 15|Set-Content (Join-Path $repo 'catalog\approved-template-additions.json') -Encoding utf8
Write-Output (Join-Path $repo 'catalog\approved-template-additions.json')
