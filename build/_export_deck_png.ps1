param([string]$Pptx, [string]$OutDir)
$ErrorActionPreference = "Stop"
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
Get-ChildItem -Path $OutDir -Filter "slide-*.png" | Remove-Item -Force
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $null
try {
  $pres = $ppt.Presentations.Open($Pptx, -1, 0, 0)
  $i = 1
  foreach ($slide in $pres.Slides) {
    $path = Join-Path $OutDir ("slide-{0:D2}.png" -f $i)
    $slide.Export($path, "PNG", 1280, 720)
    $i++
  }
  Write-Output ("SLIDES=" + $pres.Slides.Count)
} finally {
  if ($null -ne $pres) { $pres.Close() }
  $ppt.Quit()
}
