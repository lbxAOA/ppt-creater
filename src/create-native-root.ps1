param(
  [Parameter(Mandatory = $true)]
  [string]$OutputPath
)
$ErrorActionPreference = 'Stop'
$app = $null
$presentation = $null
try {
  $app = New-Object -ComObject PowerPoint.Application
  $app.Visible = -1
  $presentation = $app.Presentations.Add()
  $presentation.SaveAs([System.IO.Path]::GetFullPath($OutputPath), 24)
  $presentation.Close()
  $presentation = $null
  Write-Output ([System.IO.Path]::GetFullPath($OutputPath))
} finally {
  if ($null -ne $presentation) { try { $presentation.Close() } catch {} }
  if ($null -ne $app) { try { $app.Quit() } catch {} }
}
