# Generate a 12-page investor briefing by filling a native AI template.
# Only existing template shapes are filled where available; injected text uses
# additional editable PowerPoint textboxes on otherwise visual-only master pages.
$ErrorActionPreference='Stop'
$repo='C:\ppt-creater'
$master=Join-Path $repo 'production-template-library\ai-industry\master.pptx'
$out=Join-Path $repo 'output\embodied-world-model-investor-briefing.pptx'
New-Item -ItemType Directory -Force -Path (Split-Path $out) | Out-Null
Copy-Item $master $out -Force

$PPT_ALIGN_LEFT=1;$PPT_ALIGN_CENTER=2;$MSO_TEXT_ORIENTATION_HORIZONTAL=1
function Add-Text($slide,$text,$left,$top,$width,$height,$size=18,$color=0xFFFFFF,$bold=$false,$align=1){
  $sh=$slide.Shapes.AddTextbox($MSO_TEXT_ORIENTATION_HORIZONTAL,$left,$top,$width,$height)
  $sh.Name=('injected_{0}' -f ([guid]::NewGuid().ToString('N').Substring(0,8)))
  $tf=$sh.TextFrame;$tf.WordWrap=-1;$tf.AutoSize=0;$tf.MarginLeft=0;$tf.MarginRight=0;$tf.MarginTop=0;$tf.MarginBottom=0
  $tf.TextRange.Text=$text;$tf.TextRange.Font.NameFarEast='Microsoft YaHei';$tf.TextRange.Font.Name='Aptos';$tf.TextRange.Font.Size=$size;$tf.TextRange.Font.Bold=if($bold){-1}else{0};$tf.TextRange.Font.Color.RGB=$color;$tf.TextRange.ParagraphFormat.Alignment=$align
  return $sh
}
function Set-Text($slide,$index,$text){$slide.Shapes.Item($index).TextFrame.TextRange.Text=$text}
function Add-SourceNote($slide,$text){try{$slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text=$text}catch{}}
$app=New-Object -ComObject PowerPoint.Application;$app.Visible=-1
try{
  $p=$app.Presentations.Open($out,$false,$false,$false)
  # Native duplicate gives a 12th source slide while preserving slide-11 animation.
  [void]$p.Slides.Item(11).Duplicate()

  # 1 Cover - visual original page + editable overlay
  $s=$p.Slides.Item(1)
  Add-Text $s '具身智能世界模型：从“能看见”到“能预测、规划与闭环执行”' 70 110 750 70 28 0xFFFFFF $true $PPT_ALIGN_LEFT
  Add-Text $s '面向投资人的行业研究简报｜2026' 72 195 470 28 15 0xD9E6F2 $false $PPT_ALIGN_LEFT
  Add-Text $s '核心判断：价值将由“机器人本体”转向“数据—世界模型—策略—部署”闭环的可复制能力。' 72 405 680 38 15 0xFFFFFF $false $PPT_ALIGN_LEFT
  Add-SourceNote $s '行业研究简报。资料来源见第 12 页。'

  # 2 Investment thesis
  $s=$p.Slides.Item(2)
  Add-Text $s '投资摘要：世界模型正在把具身智能从“任务脚本”推向“可泛化的物理决策系统”' 68 55 820 45 24 0x0B1F3A $true $PPT_ALIGN_LEFT
  Add-Text $s '为什么是现在' 86 165 190 25 17 0x0B1F3A $true $PPT_ALIGN_LEFT
  Add-Text $s '多模态基础模型、仿真与机器人数据管线开始耦合；能力边界从单点抓取扩展至长程任务。' 86 205 330 80 15 0x334155 $false $PPT_ALIGN_LEFT
  Add-Text $s '投资主线' 470 165 190 25 17 0x0B1F3A $true $PPT_ALIGN_LEFT
  Add-Text $s '优先寻找掌握真实部署数据、具备高频任务闭环、并能压低验证与交付成本的团队。' 470 205 330 80 15 0x334155 $false $PPT_ALIGN_LEFT
  Add-Text $s '关键风险' 86 330 190 25 17 0x0B1F3A $true $PPT_ALIGN_LEFT
  Add-Text $s '泛化、可靠性、安全、单位经济与现场集成仍决定从演示走向订单的速度。' 86 370 330 65 15 0x334155 $false $PPT_ALIGN_LEFT
  Add-Text $s '本报告边界' 470 330 190 25 17 0x0B1F3A $true $PPT_ALIGN_LEFT
  Add-Text $s '聚焦世界模型作为多模态、目标驱动、闭环预测与规划系统；不将其狭义等同于视觉生成模型。' 470 370 350 65 15 0x334155 $false $PPT_ALIGN_LEFT

  # 3 landscape
  $s=$p.Slides.Item(3)
  Add-Text $s '产业图谱：模型、数据、仿真、机器人本体与场景交付共同构成护城河' 65 52 825 43 24 0x0B1F3A $true $PPT_ALIGN_LEFT
  Add-Text $s '上游：算力 / 传感器 / 仿真' 90 160 210 30 16 0x0B1F3A $true $PPT_ALIGN_LEFT
  Add-Text $s '中游：世界模型 / VLA / 策略' 375 160 220 30 16 0x0B1F3A $true $PPT_ALIGN_LEFT
  Add-Text $s '下游：工业、物流、家庭、医疗' 660 160 220 30 16 0x0B1F3A $true $PPT_ALIGN_LEFT
  Add-Text $s '竞争从“单机器人能力展示”转向“跨本体数据与真实工作流的复用效率”。' 105 405 700 34 17 0x334155 $true $PPT_ALIGN_CENTER

  # 4 existing slots
  $s=$p.Slides.Item(4)
  Set-Text $s 98 '世界模型的核心作用'
  Set-Text $s 99 '以可预测的环境—本体状态为中间层，连接感知、规划、控制与真实反馈。'

  # 5 world model definition
  $s=$p.Slides.Item(5)
  Add-Text $s '世界模型不是“视频生成”，而是面向行动后果的状态预测与约束推理' 70 55 820 44 24 0x0B1F3A $true $PPT_ALIGN_LEFT
  Add-Text $s '输入：视觉、触觉、语言、关节与任务状态' 95 182 280 45 16 0x0B1F3A $true $PPT_ALIGN_LEFT
  Add-Text $s '建模：对象、关系、动力学、不确定性与风险' 360 250 280 45 16 0x0B1F3A $true $PPT_ALIGN_LEFT
  Add-Text $s '输出：候选行动、预期后果、置信度与回退策略' 620 330 270 45 16 0x0B1F3A $true $PPT_ALIGN_LEFT
  Add-Text $s '投资含义：可迁移的世界表征与闭环评测，比单一场景成功率更接近长期平台价值。' 85 440 760 35 15 0x334155 $false $PPT_ALIGN_CENTER

  # 6 animated competitor comparison
  $s=$p.Slides.Item(6)
  Set-Text $s 59 '技术路径比较'
  Set-Text $s 60 '世界模型与 VLA 的分工正在收敛为“慢推理 + 快控制 + 反馈学习”。'
  Set-Text $s 61 '平台型机会来自跨本体数据、仿真验证与部署工具链，而非单一模型参数。'
  Add-SourceNote $s '参考：Google DeepMind Gemini Robotics；NVIDIA Cosmos；Figure Helix；Physical Intelligence π0。'

  # 7 animated KPI page
  $s=$p.Slides.Item(7)
  Set-Text $s 48 '投资人应追踪的五项验证指标'
  Set-Text $s 49 '不以“视频演示”替代长期、可审计的生产能力。'
  Set-Text $s 50 '01 任务泛化'
  Set-Text $s 51 '跨对象、跨环境与跨本体的成功率。'
  Set-Text $s 52 '02 数据闭环'
  Set-Text $s 53 '真实部署数据是否反哺策略迭代。'
  Set-Text $s 54 '03 单位经济'
  Set-Text $s 55 '推理、远程运维与人工兜底成本。'

  # 8 roadmap
  $s=$p.Slides.Item(8)
  Set-Text $s 107 '未来 24 个月：从技能包到可审计自治'
  Add-Text $s '阶段一：高频单任务' 85 160 190 25 16 0x0B1F3A $true $PPT_ALIGN_LEFT
  Add-Text $s '阶段二：跨工具工作流' 385 255 190 25 16 0x0B1F3A $true $PPT_ALIGN_LEFT
  Add-Text $s '阶段三：多约束自治' 670 350 190 25 16 0x0B1F3A $true $PPT_ALIGN_LEFT
  Add-Text $s '先在可控场景证明 ROI，再扩大环境变化、任务长度与安全约束。' 90 450 710 30 15 0x334155 $false $PPT_ALIGN_CENTER

  # 9 risks
  $s=$p.Slides.Item(9)
  Set-Text $s 31 '核心风险：不是模型是否会“思考”，而是系统能否稳定承担结果责任'
  Set-Text $s 32 '技术风险：长程误差累积、仿真偏差与安全边界。'
  Set-Text $s 33 '商业风险：数据采集成本、现场集成周期与人机协作责任划分。'

  # 10 recommendation
  $s=$p.Slides.Item(10)
  Set-Text $s 109 '投资建议：押注“真实部署数据 + 世界模型 + 高价值工作流”的闭环'
  Set-Text $s 110 '优先验证可重复交付能力、客户 ROI 与持续学习，而非一次性 demo。'

  # 11 closing
  $s=$p.Slides.Item(11)
  Set-Text $s 193 '世界模型将成为具身智能的“预测与决策中枢”'
  Add-Text $s '问题不在于机器人是否能完成一次任务，而在于能否在目标、风险与约束下持续做出更优行动。' 110 390 730 40 17 0x0B1F3A $true $PPT_ALIGN_CENTER

  # 12 source appendix based on duplicated native closing page
  $s=$p.Slides.Item(12)
  Set-Text $s 193 '资料来源与投资尽调提示'
  Add-Text $s '公开一手资料（截至 2026-08）' 95 125 700 30 20 0x0B1F3A $true $PPT_ALIGN_LEFT
  Add-Text $s '1. Google DeepMind｜Gemini Robotics / ER：具身推理、空间理解与低层控制协同\r2. NVIDIA｜Cosmos：世界基础模型、可控仿真与合成数据闭环\r3. Figure｜Helix：慢语义规划 + 快速连续控制的层级 VLA\r4. Physical Intelligence｜π0：跨本体、多任务数据驱动的通用策略' 100 180 760 170 15 0x334155 $false $PPT_ALIGN_LEFT
  Add-Text $s '投资尽调：所有市场规模、融资、订单、部署小时数与客户 ROI 在正式路演前应逐项替换为可追溯的一手证据。' 100 390 760 60 14 0x0B1F3A $true $PPT_ALIGN_LEFT
  Add-SourceNote $s 'https://deepmind.google/blog/gemini-robotics-brings-ai-into-the-physical-world/ | https://www.nvidia.com/en-us/ai/cosmos/ | https://www.figure.ai/news/helix | https://physicalintelligence.company/blog/pi0'

  $p.Save();$p.Close()
} finally {try{$app.Quit()}catch{};try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null}catch{}}
Write-Output $out
