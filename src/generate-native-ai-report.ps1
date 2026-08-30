$ErrorActionPreference = 'Stop'
$source = 'C:\PPT模板\01_应用场景\02_汇报与答辩\模板\101_16比9比例.pptx'
$output = 'C:\ppt-creater\output\2026-ai-industry-report-native-template.pptx'
$keep = @(1,3,4,7,14,15,16,18,22,28)

function Set-ShapeText($slide, [int]$shapeIndex, [string]$text) {
  $shape = $slide.Shapes.Item($shapeIndex)
  if (-not $shape.HasTextFrame) { throw "Slide $($slide.SlideIndex), shape $shapeIndex has no text frame." }
  $shape.TextFrame.TextRange.Text = $text
}
function Safe-Set($slide, [int]$shapeIndex, [string]$text) {
  try { Set-ShapeText $slide $shapeIndex $text } catch { Write-Warning $_.Exception.Message }
}

$app = New-Object -ComObject PowerPoint.Application
$app.Visible = -1
try {
  Copy-Item -LiteralPath $source -Destination $output -Force
  $p = $app.Presentations.Open($output, $false, $false, $false)

  # Delete all unselected pages from the copied native deck, backwards so index mapping is stable.
  for ($i=$p.Slides.Count; $i -ge 1; $i--) { if ($keep -notcontains $i) { $p.Slides.Item($i).Delete() } }

  # Slide 1: original cover. Only existing text shapes are changed.
  $s=$p.Slides.Item(1)
  Safe-Set $s 4 '2026 AI 行业'
  Safe-Set $s 6 '从模型能力竞争走向系统价值兑现'

  # Slide 2: original contents page, retained visual grid.
  $s=$p.Slides.Item(2)
  Safe-Set $s 7 '行业汇报目录'
  Safe-Set $s 9 '1';  Safe-Set $s 10 '产业价值栈'
  Safe-Set $s 12 '2'; Safe-Set $s 13 '技术与工作流'
  Safe-Set $s 15 '3'; Safe-Set $s 16 '商业化与风险'
  Safe-Set $s 18 '4'; Safe-Set $s 19 '行动建议与结论'

  # Slide 3: original section divider.
  $s=$p.Slides.Item(3)
  Safe-Set $s 5 'PART 01'
  Safe-Set $s 6 '产业价值栈与市场结构'

  # Slide 4: original animated four-person layout. Animation targets remain intact.
  $s=$p.Slides.Item(4)
  Safe-Set $s 1 '应用与工作流'
  Safe-Set $s 2 '行业 Copilot、智能体与自动化系统直接承担业务结果。'
  Safe-Set $s 3 '模型与平台'
  Safe-Set $s 4 '多模态、推理、评测与部署共同决定平台能力。'
  Safe-Set $s 5 '数据与工具链'
  Safe-Set $s 6 '数据治理、检索、观测、安全构成落地底座。'
  Safe-Set $s 7 '基础设施'
  Safe-Set $s 8 '芯片、云、存储和推理效率决定规模边界。'

  # Slide 5: original animated four-step sequence.
  $s=$p.Slides.Item(5)
  Safe-Set $s 13 '感知与上下文'
  Safe-Set $s 14 '推理与世界建模'
  Safe-Set $s 15 '规划与工具调用'
  Safe-Set $s 16 '执行与反馈闭环'
  Safe-Set $s 17 '将多模态输入、企业知识与状态转化为可执行决策。'
  Safe-Set $s 18 '预测行动后果，在不确定性下形成任务表征。'
  Safe-Set $s 19 '分解目标、调用系统、与人类和软件协同完成任务。'
  Safe-Set $s 20 '写入真实工作流，监控质量、成本并持续纠错。'

  # Slide 6: original animated feature layout.
  $s=$p.Slides.Item(6)
  Safe-Set $s 2 'AI 价值链的高价值环节'
  Safe-Set $s 3 '从能力供给到结果交付，利润池正在向最后一公里迁移。'
  Safe-Set $s 4 '模型供给'
  Safe-Set $s 8 '平台集成'
  Safe-Set $s 11 '行业应用'
  Safe-Set $s 14 '服务与治理'

  # Slide 7: original animated ratio page.
  $s=$p.Slides.Item(7)
  Safe-Set $s 5 '可验证 ROI'
  Safe-Set $s 6 '业务成果与人效改善应成为首要衡量。'
  Safe-Set $s 7 '工作流嵌入'
  Safe-Set $s 8 '日常使用深度决定产品留存和扩张。'
  Safe-Set $s 9 '数据飞轮'
  Safe-Set $s 10 '专有反馈能形成长期差异化。'
  Safe-Set $s 11 '单位经济'
  Safe-Set $s 12 '推理与交付成本必须随规模下降。'

  # Slide 8: original KPI page.
  $s=$p.Slides.Item(8)
  Safe-Set $s 5 '90 天'
  Safe-Set $s 6 '业务价值验证窗口'
  Safe-Set $s 7 '4 个'
  Safe-Set $s 8 '关键评估维度'
  Safe-Set $s 9 '可量化结果'
  Safe-Set $s 10 '任务频率、准确率、时延、人工介入'
  Safe-Set $s 11 '可复制工作流'
  Safe-Set $s 12 '连接数据、权限、工具与治理机制'
  Safe-Set $s 13 '可扩展组织能力'
  Safe-Set $s 14 '以试点为起点，形成跨团队的运营体系'

  # Slide 9: original metrics dashboard.
  $s=$p.Slides.Item(9)
  Safe-Set $s 13 '高'; Safe-Set $s 14 '中'; Safe-Set $s 15 '低'; Safe-Set $s 16 '高'
  Safe-Set $s 17 '可靠性风险'; Safe-Set $s 18 '模型幻觉、数据权限与安全边界'
  Safe-Set $s 19 '成本风险'; Safe-Set $s 20 '推理成本、交付成本与规模效率'
  Safe-Set $s 21 '组织风险'; Safe-Set $s 22 '使用习惯、流程改造与责任划分'
  Safe-Set $s 23 '治理风险'; Safe-Set $s 24 '评测、审计、回滚与合规机制'

  # Slide 10: original thank-you slide.
  $s=$p.Slides.Item(10)
  Safe-Set $s 1 '结论：AI 行业的长期机会，在于将模型能力嵌入真实数据、工作流与结果责任。'
  Safe-Set $s 6 'THANKS'

  $p.Save()
  $p.Close()
} finally {
  $app.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
Write-Output $output
