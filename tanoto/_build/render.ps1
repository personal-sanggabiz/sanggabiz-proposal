param([string]$pptx,[string]$out,[string]$pdf="")
New-Item -ItemType Directory -Force $out | Out-Null
Get-ChildItem $out -Filter *.PNG | Remove-Item -Force
$pp=New-Object -ComObject PowerPoint.Application
$pres=$pp.Presentations.Open($pptx,$true,$false,$false)
$pres.Export($out,"PNG",1600,900)
if($pdf -ne ""){ $pres.SaveAs($pdf,32) }
$pres.Close(); $pp.Quit()
"done"
