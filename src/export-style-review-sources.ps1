$ErrorActionPreference='Stop'
$outDir='C:\ppt-creater\output\style-review-sources';New-Item -ItemType Directory -Force -Path $outDir|Out-Null
$files=@(
 'C:\PPT模板\01_应用场景\02_汇报与答辩\模板\101_16比9比例.pptx',
 'C:\PPT模板\04_设计素材\04_汇报总结模板\模板\001_汇报总结模板.pptx',
 'C:\PPT模板\02_视觉风格\07_扁平化\模板\015_扁平化模板.pptx',
 'C:\PPT模板\03_主题场景\01_主题活动\模板\011_PPT模板宣传类.pptx',
 'C:\PPT模板\03_主题场景\01_主题活动\模板\003_PPT模板.pptx',
 'C:\PPT模板\01_应用场景\04_求职与竞聘\模板\039_模板组.pptx'
)
$app=New-Object -ComObject PowerPoint.Application;$app.Visible=-1
try{foreach($f in $files){$p=$app.Presentations.Open($f,$true,$false,$false);$base=[IO.Path]::GetFileNameWithoutExtension($f);$pdf=Join-Path $outDir ($base+'.pdf');$method=$p.GetType().InvokeMember('ExportAsFixedFormat',[System.Reflection.BindingFlags]::InvokeMethod,$null,$p,@($pdf,[int]2));$p.Close();Write-Output $pdf}}finally{try{$app.Quit()}catch{};try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null}catch{}}
