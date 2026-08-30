$ErrorActionPreference = 'Stop'
$input = 'C:\ppt-creater\output\2026-ai-industry-report-cn.pptx'
$preview = 'C:\ppt-creater\output\preview'
New-Item -ItemType Directory -Force -Path $preview | Out-Null
$powerPoint = New-Object -ComObject PowerPoint.Application
$powerPoint.Visible = -1
try {
  $pres = $powerPoint.Presentations.Open($input, $true, $false, $false)
  for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $dest = Join-Path $preview ('slide-{0:D2}.png' -f $i)
    $pres.Slides.Item($i).Export($dest, 'PNG', 1600, 900)
  }
  $pdf = 'C:\ppt-creater\output\2026-ai-industry-report-cn.pdf'
  $pres.SaveAs($pdf, 32)
  $pres.Close()
} finally {
  $powerPoint.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerPoint) | Out-Null
}
