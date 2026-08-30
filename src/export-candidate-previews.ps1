# Export representative pages for every candidate template family.
$ErrorActionPreference = 'Stop'
$registry = Get-Content 'C:\ppt-creater\catalog\families.json' -Raw | ConvertFrom-Json
$outRoot = 'C:\ppt-creater\candidate-previews'
New-Item -ItemType Directory -Force -Path $outRoot | Out-Null
$app = New-Object -ComObject PowerPoint.Application
$app.Visible = -1
$manifest = @()
try {
  foreach ($family in $registry.families) {
    $source = $family.source_template
    $dir = Join-Path $outRoot $family.family_id
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    try {
      $p = $app.Presentations.Open($source, $true, $false, $false)
      $count = $p.Slides.Count
      # Cover, a very early structure page, several mid-deck sample pages, and closing.
      $indices = @(
        1,
        [Math]::Min(3, $count),
        [Math]::Max(1, [Math]::Floor($count * .30)),
        [Math]::Max(1, [Math]::Floor($count * .52)),
        [Math]::Max(1, [Math]::Floor($count * .72)),
        $count
      ) | Select-Object -Unique
      $exports = @()
      foreach($index in $indices) {
        $dest = Join-Path $dir ('slide-{0:D3}.png' -f $index)
        $p.Slides.Item($index).Export($dest, 'PNG', 640, 360)
        $exports += $dest
      }
      $manifest += [ordered]@{family_id=$family.family_id;source=$source;slide_count=$count;exports=$exports;status='ok'}
      $p.Close()
    } catch {
      $manifest += [ordered]@{family_id=$family.family_id;source=$source;status='error';error=$_.Exception.Message}
      if($p){try{$p.Close()}catch{}}
    }
  }
} finally { $app.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null; [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
$manifest | ConvertTo-Json -Depth 5 | Set-Content -Path (Join-Path $outRoot 'export-manifest.json') -Encoding utf8
Write-Output (Join-Path $outRoot 'export-manifest.json')
