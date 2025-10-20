$source = "dist\test.exe"
$destination = "\\hokkaido\Daten\IT\Mu\Tools\test.exe"

Copy-Item $source -Destination $destination -Force
