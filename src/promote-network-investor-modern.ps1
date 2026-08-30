# Promotes NET-01 into a separately routed production family. It remains
# isolated from local-template families and carries an external-license review flag.
$ErrorActionPreference='Stop'
$source='C:\ppt-creater\network-supplements\NET-01\pitch-with-confidence.pptx'
$masterDir='C:\ppt-creater\network-production-library\network-investor-modern'
$validationDir='C:\ppt-creater\network-production-library\validation\network-investor-modern'
$master=Join-Path $masterDir 'master.pptx';$validation=Join-Path $validationDir 'filled-validation.pptx';$catalog='C:\ppt-creater\catalog\network-production-template-library.json'
New-Item -ItemType Directory -Force -Path $masterDir,$validationDir|Out-Null;Copy-Item $source $master -Force;Copy-Item $master $validation -Force
function DemoText($old,$n){$pool=@('投资人标题','核心判断','市场机会','产品价值','增长路径','客户验证','竞争壁垒','融资计划','团队能力','里程碑','风险控制','行动建议');$v=$pool[($n-1)%$pool.Count];return $v.Substring(0,[Math]::Min([Math]::Max(2,$old.Length),$v.Length))}
$app=New-Object -ComObject PowerPoint.Application;$app.Visible=-1
try{
 $m=$app.Presentations.Open($master,$true,$false,$false);$before=0;for($s=1;$s -le $m.Slides.Count;$s++){$before+=$m.Slides.Item($s).TimeLine.MainSequence.Count};$m.Close()
 $p=$app.Presentations.Open($validation,$false,$false,$false);$slots=@();$pages=@();$fill=0
 for($s=1;$s -le $p.Slides.Count;$s++){$page=@();for($i=1;$i -le $p.Slides.Item($s).Shapes.Count;$i++){$sh=$p.Slides.Item($s).Shapes.Item($i);$old='';$ok=$false;try{$ok=$sh.HasTextFrame -and $sh.TextFrame.HasText;if($ok){$old=$sh.TextFrame.TextRange.Text}}catch{};if($ok -and $old.Trim().Length -gt 1){$name=('slot_s{0:D2}_{1:D2}' -f $s,$i);$sh.Name=$name;$sh.TextFrame.TextRange.Text=(DemoText $old ($fill+1));$item=[ordered]@{name=$name;shape_index=$i;max_chars_cn=$old.Length;animation_locked=$false};$slots+=$item;$page+=$item;$fill++}};$pages += [ordered]@{page_index=$s;slots=$page}}
 $p.Save();$p.Close()
 $v=$app.Presentations.Open($validation,$true,$false,$false);$after=0;for($s=1;$s -le $v.Slides.Count;$s++){$after+=$v.Slides.Item($s).TimeLine.MainSequence.Count};$play=$true;try{$v.SlideShowSettings.ShowType=2;$show=$v.SlideShowSettings.Run();Start-Sleep -Milliseconds 300;$show.View.GotoSlide(1,$true);Start-Sleep -Milliseconds 100;try{$show.View.Exit()}catch{}}catch{$play=$false};try{$v.Close()}catch{}
 $data=[ordered]@{schema_version=1;overall_status=if($play -and $before -eq $after -and $fill -gt 0){'production_ready'}else{'review_required'};families=[ordered]@{'network-investor-modern'=[ordered]@{source_id='NET-01';source_pptx=$source;source_url='https://slidesmania.com/free-simple-business-pitch-free-template/';master_pptx=$master;validation_pptx=$validation;route_targets=@('ai-industry','investor-pitch');slide_count=24;semantic_slots=$slots;pages=$pages;fill_count=$fill;new_shapes=0;strict_native_only=$true;animation_effects_before=$before;animation_effects_after=$after;powerpoint_playback=$play;license_status='manual_external_review_required';status=if($play -and $before -eq $after -and $fill -gt 0){'production_ready'}else{'review_required'}}}}
 $data|ConvertTo-Json -Depth 10|Set-Content $catalog -Encoding utf8
}finally{try{$app.Quit()}catch{};try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null}catch{}}
Write-Output $catalog
