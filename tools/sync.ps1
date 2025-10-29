$source = "test\dist\MonatsberichtTeilen.exe"
$destination = "\\hokkaido\Daten\IT\Mu\Tools\MonatsberichtTeilen\MonatsberichtTeilen.exe"

Copy-Item $source -Destination $destination -Force
