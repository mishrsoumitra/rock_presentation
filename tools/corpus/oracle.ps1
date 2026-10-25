# tools/corpus/oracle.ps1  -- O2, O3 and O4 for one file. Called by the Node harness; writes oracle.json and slide PNGs.
# Runs on the Windows CI runner with Microsoft 365 installed, in an interactive session (COM automation of Office
# does not work from a service), with the corpus folder as a Trusted Location. Taken from the build plan; it has not
# been run yet (the machine this was written on is a Mac).
param([string]$In, [string]$OutDir, [int]$Width = 1920)
New-Item -ItemType Directory -Force $OutDir | Out-Null
$app = New-Object -ComObject PowerPoint.Application
$app.DisplayAlerts = 1                                  # ppAlertsNone: a repair prompt becomes an exception instead
$app.AutomationSecurity = 3                             # msoAutomationSecurityForceDisable: never run macros
$r = @{ file = $In; opened = $false; error = $null; slides = @(); shapes = @() }
try {
  $pres = $app.Presentations.Open($In, -1, 0, 0)        # ReadOnly=msoTrue, Untitled=msoFalse, WithWindow=msoFalse
  $r.opened = $true
  $h = [int]($Width * $pres.PageSetup.SlideHeight / $pres.PageSetup.SlideWidth)
  $n = 0
  foreach ($slide in $pres.Slides) {
    $n++
    $slide.Export("$OutDir\slide$n.png", "PNG", $Width, $h)
    $r.slides += @{ index = $n; id = $slide.SlideID; hidden = [bool]$slide.SlideShowTransition.Hidden }
    foreach ($shape in $slide.Shapes) {
      $entry = @{ slide = $n; id = $shape.Id; name = $shape.Name; left = $shape.Left; top = $shape.Top
                  width = $shape.Width; height = $shape.Height; rotation = $shape.Rotation; lines = @() }
      if ($shape.HasTextFrame -eq -1) {
        $tr = $shape.TextFrame2.TextRange
        $count = $tr.Lines().Count
        for ($k = 1; $k -le $count; $k++) {
          $ln = $tr.Lines($k, 1)
          $entry.lines += @{ text = $ln.Text; left = $ln.BoundLeft; top = $ln.BoundTop; width = $ln.BoundWidth; height = $ln.BoundHeight }
        }
      }
      $r.shapes += $entry
    }
  }
  $pres.SaveCopyAs("$OutDir\resaved.pptx")                 # PowerPoint's own re-save, for the O2 comparison
  $pres.Close()
} catch { $r.error = $_.Exception.Message }
finally { $app.Quit() }
$r | ConvertTo-Json -Depth 8 | Set-Content "$OutDir\oracle.json"
