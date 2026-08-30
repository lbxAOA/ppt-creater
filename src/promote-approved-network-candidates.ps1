# Promote user-approved network templates into separate, strict-native families.
# No network family is mixed with local source decks; licensing remains manually review-gated.
$ErrorActionPreference='Stop'
$repo='C:\ppt-creater';$catalog=Join-Path $repo 'catalog\network-production-template-library.json'
$families=@(
 @{id='network-company-professional';sourceId='NET-02';input='C:\ppt-creater\network-supplements\NET-02\simple-professional.pptx';targets=@('brand-company-profile','data-boardroom','generic-corporate')},
 @{id='network-career-portfolio';sourceId='NET-03';input='C:\ppt-creater\network-supplements\NET-03\my-portfolio.pptx';targets=@('career-portfolio')},
 @{id='network-wedding-amelia';sourceId='NET-04';input='C:\ppt-creater\network-supplements\NET-04\amelia-wedding.pptx';targets=@('wedding-bridal')}
)
function DemoText($old,$n){$pool=@('品牌标题','核心价值','市场机会','产品服务','团队能力','发展路径','客户案例','项目成果','重要节点','行动建议','幸福时刻','感谢见证');$v=$pool[($n-1)%$pool.Count];return $v.Substring(0,[Math]::Min([Math]::Max(2,$old.Length),$v.Length))}
function Close-App($app){try{$app.Quit()}catch{};try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null}catch{}}
$old=Get-Content $catalog -Raw|ConvertFrom-Json;$reg=[ordered]@{schema_version=1;overall_status='production_ready';families=[ordered]@{}}
foreach($prop in $old.families.psobject.Properties){$reg.families[$prop.Name]=$prop.Value}
foreach($spec in $families){
 $app=$null
 try{
  $masterDir=Join-Path $repo ('network-production-library\'+$spec.id);$validationDir=Join-Path $repo ('network-production-library\validation\'+$spec.id);New-Item -ItemType Directory -Force -Path $masterDir,$validationDir|Out-Null;$master=Join-Path $masterDir 'master.pptx';$validation=Join-Path $validationDir 'filled-validation.pptx';Copy-Item $spec.input $master -Force;Copy-Item $master $validation -Force
  $app=New-Object -ComObject PowerPoint.Application;$app.Visible=-1;$a=$app.Presentations.Open($master,$true,$false,$false);$before=0;for($s=1;$s -le $a.Slides.Count;$s++){$before+=$a.Slides.Item($s).TimeLine.MainSequence.Count};$slideCount=$a.Slides.Count;$a.Close()
  $p=$app.Presentations.Open($validation,$false,$false,$false);$slots=@();$pages=@();$fill=0
  for($s=1;$s -le $p.Slides.Count;$s++){$page=@();for($i=1;$i -le $p.Slides.Item($s).Shapes.Count;$i++){$sh=$p.Slides.Item($s).Shapes.Item($i);$oldText='';$ok=$false;try{$ok=$sh.HasTextFrame -and $sh.TextFrame.HasText;if($ok){$oldText=$sh.TextFrame.TextRange.Text}}catch{};if($ok -and $oldText.Trim().Length -gt 1){$name=('slot_s{0:D2}_{1:D2}' -f $s,$i);$sh.Name=$name;$sh.TextFrame.TextRange.Text=(DemoText $oldText ($fill+1));$it=[ordered]@{name=$name;shape_index=$i;max_chars_cn=$oldText.Length;animation_locked=($p.Slides.Item($s).TimeLine.MainSequence.Count -gt 0)};$slots+=$it;$page+=$it;$fill++}};$pages+=[ordered]@{page_index=$s;slots=$page}}
  $p.Save();$p.Close();$v=$app.Presentations.Open($validation,$true,$false,$false);$after=0;for($s=1;$s -le $v.Slides.Count;$s++){$after+=$v.Slides.Item($s).TimeLine.MainSequence.Count};$play=$true;try{$v.SlideShowSettings.ShowType=2;$show=$v.SlideShowSettings.Run();Start-Sleep -Milliseconds 300;$show.View.GotoSlide(1,$true);Start-Sleep -Milliseconds 100;try{$show.View.Exit()}catch{}}catch{$play=$false};try{$v.Close()}catch{}
  $status=if($play -and $before -eq $after -and $fill -gt 0){'production_ready'}else{'review_required'}
  $reg.families[$spec.id]=[ordered]@{source_id=$spec.sourceId;source_pptx=$spec.input;master_pptx=$master;validation_pptx=$validation;route_targets=$spec.targets;slide_count=$slideCount;semantic_slots=$slots;pages=$pages;fill_count=$fill;new_shapes=0;strict_native_only=$true;animation_effects_before=$before;animation_effects_after=$after;powerpoint_playback=$play;license_status='manual_external_review_required';status=$status}
 }finally{Close-App $app}
}
if(@($reg.families.Values|Where-Object {$_.status -ne 'production_ready'}).Count -gt 0){$reg.overall_status='review_required'}
$reg|ConvertTo-Json -Depth 12|Set-Content $catalog -Encoding utf8
Write-Output $catalog
