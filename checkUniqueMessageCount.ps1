cd "C:\Users\jsble\OneDrive\ccps\Email PIA\Final - Copy\McConkey, Kelly\To"
$files = Get-ChildItem -Recurse | Sort-Object size -Descending

$counts = $files | Group-Object length | Sort-Object Name -Descending
$counts
$count = $counts.Count
$count

