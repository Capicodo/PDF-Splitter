$source = "dist\MonatsberichtTeilen.exe"
$destination = "\\hokkaido\Daten\IT\Mu\Tools\MonatsberichtTeilenSortieren.exe"

Copy-Item $source -Destination $destination -Force
