# Strict native-template deck: only selected original slide pages and existing text shapes.
# No new shapes, no manual font/size/color overrides, no cross-family pages.
$ErrorActionPreference='Stop'
$source='C:\PPT模板\01_应用场景\02_汇报与答辩\模板\101_16比9比例.pptx'
$output='C:\ppt-creater\output\embodied-world-model-investor-briefing-strict-native.pptx'
$qaPath='C:\ppt-creater\catalog\embodied-world-model-strict-native-qa.json'
New-Item -ItemType Directory -Force -Path (Split-Path $output) | Out-Null
Copy-Item $source $output -Force

# Source indices selected for their existing information-bearing text structures.
$selected=@(1,3,4,7,9,11,14,15,16,18,22,28)
function Replace($slide,$shape,$value){
  $old=$slide.Shapes.Item($shape).TextFrame.TextRange.Text
  if($value.Length -gt [Math]::Max(8,$old.Length*2)){throw "Text exceeds conservative native capacity on slide $($slide.SlideIndex), shape $shape"}
  $slide.Shapes.Item($shape).TextFrame.TextRange.Text=$value
}
$app=New-Object -ComObject PowerPoint.Application;$app.Visible=-1
try {
  $p=$app.Presentations.Open($output,$false,$false,$false)
  $before=0;foreach($i in $selected){$before += $p.Slides.Item($i).TimeLine.MainSequence.Count}
  for($i=$p.Slides.Count;$i -ge 1;$i--){if($selected -notcontains $i){$p.Slides.Item($i).Delete()}}
  if($p.Slides.Count -ne 12){throw 'Wrong selected slide count'}

  # 1 Cover
  Replace $p.Slides.Item(1) 4 '具身智能世界模型'
  Replace $p.Slides.Item(1) 6 '投资人行业汇报｜2026'

  # 2 Agenda
  Replace $p.Slides.Item(2) 7 '报告目录'
  Replace $p.Slides.Item(2) 19 '投资判断'
  Replace $p.Slides.Item(2) 10 '世界模型闭环'
  Replace $p.Slides.Item(2) 13 '竞争与壁垒'
  Replace $p.Slides.Item(2) 16 '风险与建议'

  # 3 Section page
  Replace $p.Slides.Item(3) 6 '投资判断'

  # 4 Animated four-card thesis
  Replace $p.Slides.Item(4) 1 '为什么是现在'
  Replace $p.Slides.Item(4) 2 '多模态模型、仿真与机器人数据开始形成可训练、可验证的闭环。'
  Replace $p.Slides.Item(4) 3 '价值迁移'
  Replace $p.Slides.Item(4) 4 '壁垒从本体展示转向数据、模型、策略与部署工具链的协同。'
  Replace $p.Slides.Item(4) 5 '首选场景'
  Replace $p.Slides.Item(4) 6 '优先高频、高价值、环境可控且可获得真实反馈的工作流。'
  Replace $p.Slides.Item(4) 7 '尽调重点'
  Replace $p.Slides.Item(4) 8 '验证跨环境泛化、人工兜底率、客户 ROI 与持续学习速度。'

  # 5 Five-step native process page
  Replace $p.Slides.Item(5) 44 '状态感知'
  Replace $p.Slides.Item(5) 45 '视觉、触觉、本体与任务状态。'
  Replace $p.Slides.Item(5) 46 '世界表征'
  Replace $p.Slides.Item(5) 47 '对象、关系、动力学与不确定性。'
  Replace $p.Slides.Item(5) 48 '后果预测'
  Replace $p.Slides.Item(5) 49 '比较候选行动的未来结果。'
  Replace $p.Slides.Item(5) 50 '约束规划'
  Replace $p.Slides.Item(5) 51 '在目标、风险与资源约束下求解。'
  Replace $p.Slides.Item(5) 52 '真实反馈'
  Replace $p.Slides.Item(5) 53 '以部署数据持续校准模型与策略。'
  Replace $p.Slides.Item(5) 71 '感知'
  Replace $p.Slides.Item(5) 72 '预测'
  Replace $p.Slides.Item(5) 73 '规划'
  Replace $p.Slides.Item(5) 74 '执行'
  Replace $p.Slides.Item(5) 75 '学习'

  # 6 Four routes
  Replace $p.Slides.Item(6) 11 '传统模块化'
  Replace $p.Slides.Item(6) 12 '可控、可解释，但在长尾场景中维护成本高。'
  Replace $p.Slides.Item(6) 14 '端到端策略'
  Replace $p.Slides.Item(6) 15 '反应快，但对数据分布与失效模式较敏感。'
  Replace $p.Slides.Item(6) 17 'VLA 路线'
  Replace $p.Slides.Item(6) 18 '语言与视觉增强泛化，仍需稳定的低层执行。'
  Replace $p.Slides.Item(6) 20 '世界模型闭环'
  Replace $p.Slides.Item(6) 21 '预测后果、评估风险并以真实反馈迭代策略。'

  # 7 Animated four investment gates
  Replace $p.Slides.Item(7) 13 '数据飞轮'
  Replace $p.Slides.Item(7) 14 '部署数据能否回流。'
  Replace $p.Slides.Item(7) 15 '仿真可信度'
  Replace $p.Slides.Item(7) 16 '覆盖失败模式并加速验证。'
  Replace $p.Slides.Item(7) 17 '跨本体迁移'
  Replace $p.Slides.Item(7) 18 '模型能否跨硬件与任务复用。'
  Replace $p.Slides.Item(7) 19 '现场经济性'
  Replace $p.Slides.Item(7) 20 '推理、运维和人工兜底成本能否被客户价值覆盖。'

  # 8 Animated feature page
  Replace $p.Slides.Item(8) 3 '世界模型的四个投资价值'
  Replace $p.Slides.Item(8) 4 '预测行动后果'
  Replace $p.Slides.Item(8) 8 '降低真实试错成本'
  Replace $p.Slides.Item(8) 11 '形成可审计安全边界'
  Replace $p.Slides.Item(8) 14 '提升跨场景迁移效率'
  Replace $p.Slides.Item(8) 2 '从一次性脚本走向可持续迭代的物理智能系统。'

  # 9 Animated four-value scorecard: four native content cards plus authored metric objects.
  Replace $p.Slides.Item(9) 5 '任务泛化：跨对象、跨环境、跨本体。'
  Replace $p.Slides.Item(9) 6 '数据质量：真实反馈、标注与回流。'
  Replace $p.Slides.Item(9) 7 '系统可靠性：失败检测与回退。'
  Replace $p.Slides.Item(9) 8 '客户价值：吞吐、质量与成本。'
  Replace $p.Slides.Item(9) 9 '25%'
  Replace $p.Slides.Item(9) 10 '50%'
  Replace $p.Slides.Item(9) 11 '75%'
  Replace $p.Slides.Item(9) 12 '100%'

  # 10 Market-validation metrics
  Replace $p.Slides.Item(10) 5 '可验证部署'
  Replace $p.Slides.Item(10) 9 '可复用数据'
  Replace $p.Slides.Item(10) 11 '可扩张 ROI'
  Replace $p.Slides.Item(10) 8 '先拿到受限环境的连续运行证据，再扩大任务与环境变化。'
  Replace $p.Slides.Item(10) 10 '跨客户积累的失败案例与恢复策略，才会形成数据护城河。'
  Replace $p.Slides.Item(10) 12 '投资结论取决于单机价值、人工替代和交付周期，而非单次演示。'

  # 11 Four risk cells
  Replace $p.Slides.Item(11) 17 '长程误差'
  Replace $p.Slides.Item(11) 18 '连续任务的累计偏差需要在线检测与回退。'
  Replace $p.Slides.Item(11) 19 '仿真偏差'
  Replace $p.Slides.Item(11) 20 '必须用真实数据校准物理与传感器分布。'
  Replace $p.Slides.Item(11) 21 '交付周期'
  Replace $p.Slides.Item(11) 22 '现场集成、流程改造与安全责任影响回款速度。'
  Replace $p.Slides.Item(11) 23 '估值纪律'
  Replace $p.Slides.Item(11) 24 '以可审计部署指标与客户留存验证平台溢价。'
  Replace $p.Slides.Item(11) 13 '部署可靠性'
  Replace $p.Slides.Item(11) 14 '数据闭环'
  Replace $p.Slides.Item(11) 15 '场景经济性'
  Replace $p.Slides.Item(11) 16 '安全治理'

  # 12 Native closing / acknowledgement page
  Replace $p.Slides.Item(12) 1 '核心结论：世界模型的价值在于，让机器人在真实反馈中预测后果、规划行动并持续降低交付成本。'
  Replace $p.Slides.Item(12) 6 'THANKS'

  $p.Save();$p.Close()
  $v=$app.Presentations.Open($output,$true,$false,$false);$after=0;$shapeCount=0
  for($s=1;$s -le $v.Slides.Count;$s++){$after+=$v.Slides.Item($s).TimeLine.MainSequence.Count;$shapeCount+=$v.Slides.Item($s).Shapes.Count}
  $play=$true
  try{$v.SlideShowSettings.ShowType=2;$show=$v.SlideShowSettings.Run();Start-Sleep -Milliseconds 400;for($s=1;$s -le $v.Slides.Count;$s++){if($v.Slides.Item($s).TimeLine.MainSequence.Count -gt 0){$show.View.GotoSlide($s,$true);[void]$show.View.Next();Start-Sleep -Milliseconds 100}};try{$show.View.Exit()}catch{}}catch{$play=$false}
  try{$v.Close()}catch{}
  [ordered]@{output_pptx=$output;source_pptx=$source;slide_count=12;source_slide_indices=$selected;new_shapes=0;animation_effects_before=$before;animation_effects_after=$after;powerpoint_playback=$play;strict_native_only=$true;status=if($before -eq $after -and $play){'verified'}else{'review_required'}} | ConvertTo-Json | Set-Content $qaPath -Encoding utf8
} finally {try{$app.Quit()}catch{};try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null}catch{}}
Write-Output $output
