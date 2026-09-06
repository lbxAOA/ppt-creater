$ErrorActionPreference='Stop'
$master='C:\ppt-creater\complete-production-library\ai-industry\master.pptx'
$app=$null;$p=$null
try {
  $app=New-Object -ComObject PowerPoint.Application;$app.Visible=-1
  $p=$app.Presentations.Open($master,$false,$false,$false)
  $renamed=0
  for($s=1;$s -le $p.Slides.Count;$s++){
    for($i=1;$i -le $p.Slides.Item($s).Shapes.Count;$i++){
      $shape=$p.Slides.Item($s).Shapes.Item($i);$has=$false;$old=''
      try{$has=$shape.HasTextFrame -and $shape.TextFrame.HasText;if($has){$old=$shape.TextFrame.TextRange.Text}}catch{}
      if($has -and $old.Trim().Length -gt 1){
        $shape.Name=('slot_s{0:D2}_{1:D2}' -f $s,$i);$renamed++
      }
    }
  }
  $slideCount=$p.Slides.Count
  $p.Save()
  $p.Close();$p=$null
  Write-Output ([ordered]@{master_pptx=$master;slide_count=$slideCount;renamed_slots=$renamed} | ConvertTo-Json -Compress)
} finally {
  if($null -ne $p){try{$p.Close()}catch{}}
  if($null -ne $app){try{$app.Quit()}catch{};try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null}catch{}}
}
