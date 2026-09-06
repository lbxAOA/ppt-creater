param(
  [Parameter(Mandatory = $true)] [string]$SourcePath,
  [Parameter(Mandatory = $true)] [string]$OutputPath,
  [Parameter(Mandatory = $true)] [string]$PlanPath,
  [Parameter(Mandatory = $true)] [string]$ReportPath,
  [string]$BeforePreviewReportPath,
  [string]$AfterPreviewReportPath
)
$ErrorActionPreference = 'Stop'

function Get-EffectCount($presentation) {
  $count = 0
  for ($slideIndex = 1; $slideIndex -le $presentation.Slides.Count; $slideIndex++) {
    $count += $presentation.Slides.Item($slideIndex).TimeLine.MainSequence.Count
  }
  return $count
}

function Get-ShapeCounts($presentation) {
  $counts = @()
  for ($slideIndex = 1; $slideIndex -le $presentation.Slides.Count; $slideIndex++) {
    $counts += $presentation.Slides.Item($slideIndex).Shapes.Count
  }
  return $counts
}

function Get-TimelineSignatures($presentation) {
  $slides = @()
  for ($slideIndex = 1; $slideIndex -le $presentation.Slides.Count; $slideIndex++) {
    $effects = @()
    $sequence = $presentation.Slides.Item($slideIndex).TimeLine.MainSequence
    for ($effectIndex = 1; $effectIndex -le $sequence.Count; $effectIndex++) {
      $effect = $sequence.Item($effectIndex)
      $target = ''
      $effectType = ''
      $trigger = ''
      $duration = ''
      $delay = ''
      try { $target = $effect.Shape.Name } catch {}
      try { $effectType = [string]$effect.EffectInformation.Type } catch {}
      try { $trigger = [string]$effect.Timing.TriggerType } catch {}
      try { $duration = [string]$effect.Timing.Duration } catch {}
      try { $delay = [string]$effect.Timing.TriggerDelayTime } catch {}
      $effects += "$target|$effectType|$trigger|$duration|$delay"
    }
    $slides += ,$effects
  }
  return ,$slides
}

function Get-SlideTimelineSignature($slide) {
  $effects = @()
  $sequence = $slide.TimeLine.MainSequence
  for ($effectIndex = 1; $effectIndex -le $sequence.Count; $effectIndex++) {
    $effect = $sequence.Item($effectIndex)
    $target = ''
    $effectType = ''
    $trigger = ''
    $duration = ''
    $delay = ''
    try { $target = $effect.Shape.Name } catch {}
    try { $effectType = [string]$effect.EffectInformation.Type } catch {}
    try { $trigger = [string]$effect.Timing.TriggerType } catch {}
    try { $duration = [string]$effect.Timing.Duration } catch {}
    try { $delay = [string]$effect.Timing.TriggerDelayTime } catch {}
    $effects += "$target|$effectType|$trigger|$duration|$delay"
  }
  return ,$effects
}

function Close-ComObject($object) {
  if ($null -ne $object) {
    try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($object) | Out-Null } catch {}
  }
}

$result = [ordered]@{
  source_pptx = [System.IO.Path]::GetFullPath($SourcePath)
  output_pptx = [System.IO.Path]::GetFullPath($OutputPath)
  plan_json = [System.IO.Path]::GetFullPath($PlanPath)
  normalized_with_powerpoint = $false
  source_master_slide_count = $null
  source_slide_count = $null
  output_slide_count = $null
  source_shape_counts = @()
  output_shape_counts = @()
  new_shapes = $null
  animation_effects_before = $null
  animation_effects_after = $null
  animation_style = 'preserve_original_native_timeline'
  slot_fill_match = $false
  powerpoint_available = $false
  powerpoint_opened = $false
  powerpoint_playback = $false
  before_preview = $null
  after_preview = $null
  preview_gates = $false
  status = 'review_required'
  error = $null
}
$app = $null
$source = $null
$output = $null
$check = $null
$show = $null
$temp = "$($result.output_pptx).normalized.tmp.pptx"
try {
  $plan = Get-Content -LiteralPath $result.plan_json -Raw | ConvertFrom-Json
  if ($BeforePreviewReportPath) {
    $result.before_preview = Get-Content -LiteralPath $BeforePreviewReportPath -Raw | ConvertFrom-Json
  }
  if ($AfterPreviewReportPath) {
    $result.after_preview = Get-Content -LiteralPath $AfterPreviewReportPath -Raw | ConvertFrom-Json
  }
  $result.preview_gates = (
    $null -ne $result.before_preview -and
    $null -ne $result.after_preview -and
    [string]$result.before_preview.status -eq 'verified' -and
    [string]$result.after_preview.status -eq 'verified' -and
    [int]$result.after_preview.exported_count -eq [int]$result.after_preview.slide_count -and
    [int]$result.after_preview.text_overflow -eq 0 -and
    [int]$result.after_preview.overlap_count -eq 0
  )
  $app = New-Object -ComObject PowerPoint.Application
  $app.Visible = -1
  $app.WindowState = if ($env:PPT_CREATER_POWERPOINT_VISIBLE -eq '1') { 1 } else { 2 }
  $app.DisplayAlerts = if ($env:PPT_CREATER_POWERPOINT_DISPLAY_ALERTS -eq '1') { 2 } else { 1 }
  $result.powerpoint_available = $true

  $source = $app.Presentations.Open($result.source_pptx, $true, $true, $false)
  $result.source_master_slide_count = $source.Slides.Count
  $selectedSourceSlides = @($plan.pages | ForEach-Object { $source.Slides.Item([int]$_.page_index) })
  $result.source_slide_count = $selectedSourceSlides.Count
  $result.source_shape_counts = @($selectedSourceSlides | ForEach-Object { $_.Shapes.Count })
  $result.animation_effects_before = [int](@($selectedSourceSlides | ForEach-Object { $_.TimeLine.MainSequence.Count } | Measure-Object -Sum).Sum)
  $sourceTimeline = @()
  foreach ($sourceSlide in $selectedSourceSlides) {
    $sourceTimeline += (Get-SlideTimelineSignature $sourceSlide | ConvertTo-Json -Compress)
  }
  $source.Close(); $source = $null

  if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Force }
  $output = $app.Presentations.Open($result.output_pptx, $false, $false, $false)
  $output.SaveAs($temp, 24)
  $output.Close(); $output = $null
  Move-Item -LiteralPath $temp -Destination $result.output_pptx -Force
  $result.normalized_with_powerpoint = $true

  $check = $app.Presentations.Open($result.output_pptx, $true, $false, $true)
  $result.powerpoint_opened = $true
  $result.output_slide_count = $check.Slides.Count
  $result.output_shape_counts = @($check.Slides | ForEach-Object { $_.Shapes.Count })
  $result.animation_effects_after = @($check.Slides | ForEach-Object { $_.TimeLine.MainSequence.Count } | Measure-Object -Sum).Sum
  $outputTimeline = @($check.Slides | ForEach-Object { (Get-SlideTimelineSignature $_ | ConvertTo-Json -Compress) })
  $result.timeline_signature_match = ($result.animation_effects_before -eq $result.animation_effects_after -and ($sourceTimeline | ConvertTo-Json -Compress) -eq ($outputTimeline | ConvertTo-Json -Compress))

  $shapeCountsMatch = (($result.source_shape_counts | ConvertTo-Json -Compress) -eq ($result.output_shape_counts | ConvertTo-Json -Compress))
  $slotChecks = @()
  foreach ($page in @($plan.pages)) {
    $outputIndex = if ($null -ne $page.index) { [int]$page.index } else { [int]$page.page_index }
    $slide = $check.Slides.Item($outputIndex)
    $expectedContent = if ($null -ne $page.content_final) { $page.content_final } else { $page.content }
    foreach ($property in $expectedContent.psobject.Properties) {
      $shape = $null
      $found = $false
      $actual = ''
      try {
        $shape = $slide.Shapes.Item([string]$property.Name)
        $found = $true
        $actual = [string]$shape.TextFrame.TextRange.Text
      } catch {}
      $slotChecks += ($found -and $actual.Trim() -eq [string]$property.Value)
    }
  }
  $result.slot_fill_match = @($slotChecks | Where-Object { -not $_ }).Count -eq 0

  try {
    $check.SlideShowSettings.ShowType = 2
    $show = $check.SlideShowSettings.Run()
    if ($null -eq $show -or $null -eq $show.View) { throw 'PowerPoint did not start a slide-show window' }
    Start-Sleep -Milliseconds 500
    for ($slideIndex = 1; $slideIndex -le $check.Slides.Count; $slideIndex++) {
      if ($check.Slides.Item($slideIndex).TimeLine.MainSequence.Count -gt 0) {
        $show.View.GotoSlide($slideIndex, $true)
        [void]$show.View.Next()
        Start-Sleep -Milliseconds 100
      }
    }
    $result.powerpoint_playback = $true
  } finally {
    if ($null -ne $show) { try { $show.View.Exit() } catch {} }
  }

  $result.new_shapes = if ($shapeCountsMatch) { 0 } else { 1 }
  if ($result.powerpoint_opened -and $result.normalized_with_powerpoint -and $result.output_slide_count -eq $result.source_slide_count -and $result.new_shapes -eq 0 -and $result.slot_fill_match -and $result.timeline_signature_match -and $result.powerpoint_playback -and $result.preview_gates) {
    $result.status = 'verified'
  }
} catch {
  $result.error = $_.Exception.Message
} finally {
  if ($null -ne $show) { try { $show.View.Exit() } catch {} }
  if ($null -ne $check) { try { $check.Close() } catch {} }
  if ($null -ne $output) { try { $output.Close() } catch {} }
  if ($null -ne $source) { try { $source.Close() } catch {} }
  if ($null -ne $app) { try { $app.Quit() } catch {} }
  Close-ComObject $check; Close-ComObject $output; Close-ComObject $source; Close-ComObject $app
  if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Force }
}
$result | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $ReportPath -Encoding UTF8
$result | ConvertTo-Json -Compress
if ($result.status -ne 'verified') { exit 1 }
