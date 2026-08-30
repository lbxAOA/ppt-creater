# Correct typography by inheriting formatting from native text shapes in the same template family.
# No manually selected fonts, sizes, colors, or bold rules are applied here.
$ErrorActionPreference='Stop'
$source='C:\ppt-creater\output\embodied-world-model-investor-briefing.pptx'
$output='C:\ppt-creater\output\embodied-world-model-investor-briefing-native-typography.pptx'
Copy-Item $source $output -Force
$app=New-Object -ComObject PowerPoint.Application;$app.Visible=-1
try {
  $p=$app.Presentations.Open($output,$false,$false,$false)
  # Exact native reference styles: slide 6 is the family’s white text treatment;
  # slide 7 is the family’s dark text treatment. PickUp/Apply carries paragraph,
  # run, font, color, spacing and textbox treatment as authored in the source deck.
  $white=$p.Slides.Item(6).Shapes.Item(59)
  $dark=$p.Slides.Item(7).Shapes.Item(49)
  foreach($slideIndex in 1..$p.Slides.Count) {
    $slide=$p.Slides.Item($slideIndex)
    $useWhite=($slideIndex -eq 1)
    for($i=1;$i -le $slide.Shapes.Count;$i++) {
      $shape=$slide.Shapes.Item($i)
      if($shape.Name.StartsWith('injected_')) {
        if($useWhite){$white.PickUp()}else{$dark.PickUp()}
        $shape.Apply()
        # Preserve the report’s intended alignment, but inherit all typography.
        if($slideIndex -in @(3,5,8,11)){$shape.TextFrame.TextRange.ParagraphFormat.Alignment=2}
      }
    }
  }
  $p.Save();$p.Close()
} finally {try{$app.Quit()}catch{};try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)|Out-Null}catch{}}
Write-Output $output
