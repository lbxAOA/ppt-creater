$repo='C:\ppt-creater'
$out=Join-Path $repo 'catalog\production-ready-template-library.json'
$root=Join-Path $repo 'production-template-library\validation'
$targets=[ordered]@{}
Get-ChildItem $root -Directory | ForEach-Object {
  $path=Join-Path $_.FullName 'validation.json'
  if(Test-Path $path){$targets[$_.Name]=Get-Content $path -Raw | ConvertFrom-Json}
}
$overall='production_ready'
foreach($t in $targets.Values){if($t.fill_status -ne 'verified' -or $t.playback_status -ne 'verified'){$overall='review_required'}}
[ordered]@{schema_version=1;overall_status=$overall;generated_at=(Get-Date).ToString('s');targets=$targets} | ConvertTo-Json -Depth 15 | Set-Content -Path $out -Encoding utf8
Write-Output $out
