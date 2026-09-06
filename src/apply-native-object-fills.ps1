param(
  [Parameter(Mandatory = $true)] [string]$InputPath,
  [Parameter(Mandatory = $true)] [string]$PlanPath
)
$ErrorActionPreference = 'Stop'

function Get-ShapeByName($shapes, [string]$name) {
  for ($i = 1; $i -le $shapes.Count; $i++) {
    $shape = $shapes.Item($i)
    try { if ([string]$shape.Name -eq $name) { return $shape } } catch {}
    try {
      if ([int]$shape.Type -eq 6) {
        $nested = Get-ShapeByName $shape.GroupItems $name
        if ($null -ne $nested) { return $nested }
      }
    } catch {}
  }
  return $null
}

function Set-ShapeText($shape, [string]$value) {
  if ($null -eq $shape) { throw 'Native object shape was not found.' }
  try {
    if ($shape.HasTextFrame) {
      $shape.TextFrame2.TextRange.Text = $value
      return
    }
  } catch {}
  try {
    $shape.TextFrame.TextRange.Text = $value
    return
  } catch {
    throw "Shape '$($shape.Name)' does not expose a text frame."
  }
}

function Set-TableCell($shape, $locator, [string]$value) {
  $row = [int]$locator.row
  $column = [int]$locator.column
  if ($row -lt 1 -or $column -lt 1) { throw 'table_cell locator requires positive row and column.' }
  Set-ShapeText $shape.Table.Cell($row, $column).Shape $value
}

function Get-SmartArtNode($nodes, [int]$index) {
  if ($index -lt 1 -or $index -gt $nodes.Count) { throw "SmartArt node index out of range: $index" }
  return $nodes.Item($index)
}

function Apply-ObjectFill($slide, $fill) {
  $name = [string]$fill.name
  $kind = [string]$fill.kind
  $shape = Get-ShapeByName $slide.Shapes $name
  if ($null -eq $shape) { throw "Native object '$name' was not found on slide $($slide.SlideIndex)." }
  $locator = $fill.locator
  $value = if ($null -eq $fill.value) { '' } else { [string]$fill.value }

  switch ($kind) {
    'text' { Set-ShapeText $shape $value }
    'group_text' { Set-ShapeText $shape $value }
    'table_cell' { Set-TableCell $shape $locator $value }
    'smartart_text' {
      if ($null -eq $locator -or $null -eq $locator.node_index) { throw "smartart_text '$name' requires locator.node_index." }
      $node = Get-SmartArtNode $shape.SmartArt.AllNodes ([int]$locator.node_index)
      $node.TextFrame2.TextRange.Text = $value
    }
    'chart_title' {
      if (-not $shape.HasChart) { throw "chart_title '$name' target is not a chart." }
      $shape.Chart.HasTitle = $true
      $shape.Chart.ChartTitle.Text = $value
    }
    'chart_label' {
      if ($null -eq $locator -or $null -eq $locator.series -or $null -eq $locator.point) { throw "chart_label '$name' requires locator.series and locator.point." }
      $shape.Chart.SeriesCollection.Item([int]$locator.series).Points.Item([int]$locator.point).DataLabel.Text = $value
    }
    'chart_data' {
      if ($null -eq $locator -or $null -eq $locator.series -or $null -eq $locator.point) { throw "chart_data '$name' requires locator.series and locator.point." }
      $shape.Chart.SeriesCollection.Item([int]$locator.series).Values([int]$locator.point) = $fill.value
    }
    'image_fill' {
      if (-not (Test-Path -LiteralPath $value -PathType Leaf)) { throw "Image fill '$name' file does not exist: $value" }
      $shape.Fill.UserPicture($value)
    }
    'image_placeholder' {
      if (-not (Test-Path -LiteralPath $value -PathType Leaf)) { throw "Image placeholder '$name' file does not exist: $value" }
      $shape.Fill.UserPicture($value)
    }
    'shape_fill' {
      $shape.Fill.ForeColor.RGB = [int]$fill.value
    }
    default { throw "Unsupported native object fill: $name, kind=$kind" }
  }
}

$app = $null
$presentation = $null
$result = [ordered]@{ status = 'failed'; input_pptx = [System.IO.Path]::GetFullPath($InputPath); applied = 0; error = $null }
try {
  $plan = Get-Content -LiteralPath $PlanPath -Raw | ConvertFrom-Json
  $app = New-Object -ComObject PowerPoint.Application
  $app.Visible = -1
  $app.WindowState = if ($env:PPT_CREATER_POWERPOINT_VISIBLE -eq '1') { 1 } else { 2 }
  $app.DisplayAlerts = if ($env:PPT_CREATER_POWERPOINT_DISPLAY_ALERTS -eq '1') { 2 } else { 1 }
  $presentation = $app.Presentations.Open($result.input_pptx, $false, $false, $false)
  foreach ($page in @($plan.pages)) {
    $slide = $presentation.Slides.Item([int]$page.index)
    foreach ($fill in @($page.object_fills)) {
      Apply-ObjectFill $slide $fill
      $result.applied++
    }
  }
  $presentation.Save()
  $result.status = 'verified'
} catch {
  $result.error = $_.Exception.Message
} finally {
  if ($null -ne $presentation) { try { $presentation.Close() } catch {} }
  if ($null -ne $app) { try { $app.Quit() } catch {} }
}
$result | ConvertTo-Json -Compress
if ($result.status -ne 'verified') { exit 1 }
