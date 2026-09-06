param(
  [Parameter(Mandatory = $true)]
  [string]$InputPath,
  [string]$ReportPath
)
$ErrorActionPreference = 'Stop'

function Normalize-ZipPath([string]$Base, [string]$Target) {
  $parts = New-Object System.Collections.Generic.List[string]
  if ($Base) {
    foreach ($part in $Base.Trim('/').Split('/')) {
      if ($part) { [void]$parts.Add($part) }
    }
  }
  foreach ($part in $Target.Replace('\', '/').Split('/')) {
    if (-not $part -or $part -eq '.') { continue }
    if ($part -eq '..') {
      if ($parts.Count -gt 0) { $parts.RemoveAt($parts.Count - 1) }
    } else {
      [void]$parts.Add($part)
    }
  }
  return ($parts -join '/')
}

function Read-ZipXml($Entry) {
  $reader = New-Object System.IO.StreamReader($Entry.Open())
  try { return [xml]$reader.ReadToEnd() }
  finally { $reader.Dispose() }
}

$path = [System.IO.Path]::GetFullPath($InputPath)
$result = [ordered]@{
  input_pptx = $path
  package_exists = Test-Path -LiteralPath $path -PathType Leaf
  zip_valid = $false
  xml_valid = $false
  relationships_valid = $false
  powerpoint_available = $false
  opened = $false
  slide_count = $null
  error = $null
}
$app = $null
$presentation = $null
$archive = $null
try {
  if (-not $result.package_exists) { throw "PPTX does not exist: $path" }
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $archive = [System.IO.Compression.ZipFile]::OpenRead($path)
  $entries = @{}
  foreach ($entry in $archive.Entries) {
    $entries[$entry.FullName.Replace('\', '/')] = $entry
  }
  foreach ($name in @('[Content_Types].xml', '_rels/.rels', 'ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels')) {
    if (-not $entries.ContainsKey($name)) { throw "Missing required package part: $name" }
  }
  $result.zip_valid = $true

  $documents = @{}
  foreach ($entry in $archive.Entries | Where-Object { $_.FullName -match '\.(xml|rels)$' }) {
    $documents[$entry.FullName.Replace('\', '/')] = Read-ZipXml $entry
  }
  $result.xml_valid = $true

  foreach ($relName in $documents.Keys | Where-Object { $_ -match '\.rels$' }) {
    $relDocument = $documents[$relName]
    $sourcePart = if ($relName -eq '_rels/.rels') { '' } else {
      ($relName -replace '/_rels/', '/') -replace '\.rels$', ''
    }
    $sourceDir = if ($sourcePart) { [System.IO.Path]::GetDirectoryName($sourcePart).Replace('\', '/') } else { '' }
    $ids = @{}
    foreach ($relationship in $relDocument.Relationships.Relationship) {
      $id = [string]$relationship.Id
      if ($ids.ContainsKey($id)) { throw "Duplicate relationship Id $id in $relName" }
      $ids[$id] = $true
      if ([string]$relationship.TargetMode -eq 'External') { continue }
      $resolved = Normalize-ZipPath $sourceDir ([string]$relationship.Target)
      if (-not $entries.ContainsKey($resolved)) { throw "Broken relationship in ${relName}: ${resolved}" }
    }
  }
  $result.relationships_valid = $true

  $presentationXml = $documents['ppt/presentation.xml']
  $slideIds = @($presentationXml.presentation.sldIdLst.sldId)
  if ($slideIds.Count -eq 0) { throw 'Presentation contains no slides.' }
  $uniqueIds = @($slideIds | ForEach-Object { [string]$_.id } | Select-Object -Unique)
  if ($uniqueIds.Count -ne $slideIds.Count) { throw 'Presentation contains duplicate slide IDs.' }

  $app = New-Object -ComObject PowerPoint.Application
  $result.powerpoint_available = $true
  $app.Visible = -1
  $app.WindowState = if ($env:PPT_CREATER_POWERPOINT_VISIBLE -eq '1') { 1 } else { 2 }
  $app.DisplayAlerts = if ($env:PPT_CREATER_POWERPOINT_DISPLAY_ALERTS -eq '1') { 2 } else { 1 }
  $presentation = $app.Presentations.Open($path, $true, $false, $false)
  $result.opened = $true
  $result.slide_count = $presentation.Slides.Count
} catch {
  $result.error = $_.Exception.Message
} finally {
  if ($null -ne $presentation) { try { $presentation.Close() } catch {} }
  if ($null -ne $archive) { $archive.Dispose() }
  if ($null -ne $app) { try { $app.Quit() } catch {} }
}

if (-not $ReportPath) { $ReportPath = [System.IO.Path]::ChangeExtension($path, '.powerpoint-qa.json') }
$result | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $ReportPath -Encoding UTF8
$result | ConvertTo-Json -Compress
if (-not $result.opened) { exit 1 }
