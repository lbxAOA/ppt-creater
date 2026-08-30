$ErrorActionPreference = 'Stop'
$root = 'C:\PPT模板'
$outRoot = 'C:\ppt-creater\legacy-converted'
$manifestPath = 'C:\ppt-creater\catalog\legacy-conversion.jsonl'
New-Item -ItemType Directory -Force -Path $outRoot | Out-Null
$files = Get-ChildItem -Path $root -Recurse -File | Where-Object { $_.Extension -ieq '.ppt' }
$existing = @{}
if (Test-Path $manifestPath) {
  Get-Content $manifestPath | Where-Object { $_.Trim() } | ForEach-Object {
    try { $r = $_ | ConvertFrom-Json; if ($r.status -eq 'ok') { $existing[$r.source] = $true } } catch {}
  }
}
$pending = @($files | Where-Object { -not $existing.ContainsKey($_.FullName) })
$writer = New-Object System.IO.StreamWriter($manifestPath, $true, [System.Text.UTF8Encoding]::new($false))
$powerPoint = New-Object -ComObject PowerPoint.Application
$powerPoint.Visible = -1
try {
  $i = 0
  foreach ($file in $pending) {
    $i++
    $relative = $file.FullName.Substring($root.Length).TrimStart('\')
    $target = Join-Path $outRoot ([System.IO.Path]::ChangeExtension($relative, '.pptx'))
    New-Item -ItemType Directory -Force -Path ([System.IO.Path]::GetDirectoryName($target)) | Out-Null
    $result = [ordered]@{ source=$file.FullName; relative_path=$relative; target=$target; status='pending' }
    try {
      $pres = $powerPoint.Presentations.Open($file.FullName, $true, $false, $false)
      $pres.SaveAs($target, 24)
      $pres.Close()
      $result.status = 'ok'
    } catch {
      $result.status = 'error'
      $result.error = $_.Exception.Message
      if ($pres) { try { $pres.Close() } catch {} }
    }
    $writer.WriteLine(($result | ConvertTo-Json -Compress))
    $writer.Flush()
    if ($i % 10 -eq 0) { Write-Output ("{`"progress`":$i,`"total`":$($pending.Count)}") }
  }
} finally {
  $writer.Close()
  $powerPoint.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerPoint) | Out-Null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
Write-Output ("{`"complete`":true,`"total`":$($pending.Count)}")
