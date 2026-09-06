param(
  [Parameter(Mandatory = $true)] [string]$InputPath,
  [Parameter(Mandatory = $true)] [string]$PreviewDirectory,
  [Parameter(Mandatory = $true)] [string]$ReportPath,
  [switch]$AllowReview
)
$ErrorActionPreference = 'Stop'

function Get-TextShapes($shapes, [string]$parentPath = '', [double]$offsetLeft = 0, [double]$offsetTop = 0) {
  $items = @()
  for ($i = 1; $i -le $shapes.Count; $i++) {
    $shape = $shapes.Item($i)
    $shapePath = if ($parentPath) { "$parentPath/$i" } else { [string]$i }
    $shapeLeft = $offsetLeft + [double]$shape.Left
    $shapeTop = $offsetTop + [double]$shape.Top
    $isGroup = $false
    try { $isGroup = ([int]$shape.Type -eq 6) } catch {}
    if ($isGroup) {
      try { $items += Get-TextShapes $shape.GroupItems $shapePath $shapeLeft $shapeTop } catch {}
      continue
    }
    $hasText = $false
    $text = ''
    try {
      $hasText = $shape.HasTextFrame -and $shape.TextFrame.HasText
      if ($hasText) { $text = [string]$shape.TextFrame.TextRange.Text }
    } catch {}
    $text = $text.Trim()
    # Only registered production slots are layout content. Unregistered text
    # (watermarks, chart callouts, and decorative glyphs) belongs to the
    # template artwork and is not a generated text collision candidate.
    if (-not $hasText -or -not $text -or ([string]$shape.Name) -notlike 'slot_*') { continue }
    $textLeft = $shapeLeft
    $textTop = $shapeTop
    $textRight = $shapeLeft + [double]$shape.Width
    $textBottom = $shapeTop + [double]$shape.Height
    try {
      $range = $shape.TextFrame2.TextRange
      $textLeft = $shapeLeft + [double]$range.BoundLeft
      $textTop = $shapeTop + [double]$range.BoundTop
      $textRight = $textLeft + [double]$range.BoundWidth
      $textBottom = $textTop + [double]$range.BoundHeight
    } catch {}
    $items += [pscustomobject]@{
      shape = $shape
      path = $shapePath
      text = $text
      left = $textLeft
      top = $textTop
      right = $textRight
      bottom = $textBottom
    }
  }
  return $items
}

function Get-TextMetrics($item) {
  $overflow = $false
  $fontSize = 0.0
  try {
    $boundHeight = [double]$item.shape.TextFrame.TextRange.BoundHeight
    $height = [double]$item.shape.Height
    $overflow = $boundHeight -gt ($height + 2)
  } catch {}
  try { $fontSize = [double]$item.shape.TextFrame.TextRange.Font.Size } catch {}
  return [pscustomobject]@{ overflow = $overflow; font_size = $fontSize }
}

function Get-BlockingOverlapCount($items) {
  $count = 0
  for ($i = 0; $i -lt $items.Count; $i++) {
    for ($j = $i + 1; $j -lt $items.Count; $j++) {
      $a = $items[$i]; $b = $items[$j]
      $width = [Math]::Max(0, [Math]::Min($a.right, $b.right) - [Math]::Max($a.left, $b.left))
      $height = [Math]::Max(0, [Math]::Min($a.bottom, $b.bottom) - [Math]::Max($a.top, $b.top))
      $intersection = $width * $height
      if ($intersection -le 0) { continue }
      $areaA = [Math]::Max(1, ($a.right - $a.left) * ($a.bottom - $a.top))
      $areaB = [Math]::Max(1, ($b.right - $b.left) * ($b.bottom - $b.top))
      $smallerArea = [Math]::Min($areaA, $areaB)
      $largerArea = [Math]::Max($areaA, $areaB)
      # Flag only material collision in both boxes; a large paragraph box
      # touching a small label at its edge is not a blocking overlap.
      if (($intersection / $smallerArea) -ge 0.15 -and ($intersection / $largerArea) -ge 0.05) { $count++ }
    }
  }
  return $count
}

$result = [ordered]@{
  input_pptx = [System.IO.Path]::GetFullPath($InputPath)
  preview_directory = [System.IO.Path]::GetFullPath($PreviewDirectory)
  slide_count = $null
  exported_count = 0
  failed_count = 0
  files = @()
  slides = @()
  all_pages_previewed = $false
  text_overflow = 0
  overlap_count = 0
  unreadable_count = 0
  status = 'review_required'
  error = $null
}
$app = $null
$presentation = $null
try {
  if (Test-Path -LiteralPath $result.preview_directory) { Remove-Item -LiteralPath $result.preview_directory -Recurse -Force }
  New-Item -ItemType Directory -Path $result.preview_directory -Force | Out-Null
  $app = New-Object -ComObject PowerPoint.Application
  $app.Visible = -1
  $app.WindowState = if ($env:PPT_CREATER_POWERPOINT_VISIBLE -eq '1') { 1 } else { 2 }
  $app.DisplayAlerts = if ($env:PPT_CREATER_POWERPOINT_DISPLAY_ALERTS -eq '1') { 2 } else { 1 }
  $presentation = $app.Presentations.Open($result.input_pptx, $true, $true, $false)
  $result.slide_count = $presentation.Slides.Count
  for ($slideIndex = 1; $slideIndex -le $presentation.Slides.Count; $slideIndex++) {
    $slide = $presentation.Slides.Item($slideIndex)
    $png = Join-Path $result.preview_directory ('slide-{0:D3}.png' -f $slideIndex)
    $slideResult = [ordered]@{
      slide_index = $slideIndex
      preview_path = $png
      exported = $false
      text_overflow = 0
      overlap_count = 0
      unreadable_count = 0
      error = $null
    }
    try {
      $slide.Export($png, 'PNG', 1600, 900)
      if (-not (Test-Path -LiteralPath $png) -or (Get-Item -LiteralPath $png).Length -le 0) { throw 'PowerPoint exported an empty preview.' }
      $slideResult.exported = $true
      $result.exported_count++
    } catch {
      $result.failed_count++
      $slideResult.error = $_.Exception.Message
    }

    $textItems = @(Get-TextShapes $slide.Shapes)
    foreach ($item in $textItems) {
      $metrics = Get-TextMetrics $item
      if ($metrics.overflow) { $slideResult.text_overflow++; $result.text_overflow++ }
      if ($metrics.font_size -gt 0 -and $metrics.font_size -lt 8) { $slideResult.unreadable_count++; $result.unreadable_count++ }
    }
    $slideResult.overlap_count = Get-BlockingOverlapCount $textItems
    $result.overlap_count += $slideResult.overlap_count
    $result.slides += [pscustomobject]$slideResult
  }
  $result.files = @($result.slides | Where-Object { $_.exported } | ForEach-Object { $_.preview_path })
  $result.all_pages_previewed = ($result.slide_count -gt 0 -and $result.exported_count -eq $result.slide_count -and $result.failed_count -eq 0)
  if ($result.all_pages_previewed -and $result.text_overflow -eq 0 -and $result.overlap_count -eq 0 -and $result.unreadable_count -eq 0) {
    $result.status = 'verified'
  }
} catch {
  $result.error = $_.Exception.Message
} finally {
  if ($null -ne $presentation) { try { $presentation.Close() } catch {} }
  if ($null -ne $app) { try { $app.Quit() } catch {} }
}
$result | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $ReportPath -Encoding UTF8
$result | ConvertTo-Json -Depth 10 -Compress
  if ($result.status -ne 'verified' -and -not $AllowReview) { exit 1 }
