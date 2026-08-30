$input='C:\ppt-creater\production-template-library\ai-industry\master.pptx'
$app=New-Object -ComObject PowerPoint.Application;$app.Visible=-1
try {
 $p=$app.Presentations.Open($input,$true,$false,$false)
 foreach($spec in @(@(6,59),@(6,60),@(6,61),@(7,48),@(7,49),@(11,193))){
  $sh=$p.Slides.Item($spec[0]).Shapes.Item($spec[1]);$f=$sh.TextFrame.TextRange.Font
  [pscustomobject]@{slide=$spec[0];shape=$spec[1];name=$sh.Name;text=$sh.TextFrame.TextRange.Text;font=$f.Name;fontFE=$f.NameFarEast;size=$f.Size;color=$f.Color.RGB;bold=$f.Bold} | ConvertTo-Json -Compress
 }
 $p.Close()
} finally {try{$app.Quit()}catch{};try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null}catch{}}
