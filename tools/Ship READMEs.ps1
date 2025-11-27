$source1 = ".\dist\README.html"
$source2 = ".\dist\README.md"

$destination = "\\hokkaido\Daten\IT\Mu\Tools\Monatsbericht Automat\"

Copy-Item $source1, $source2 -Destination $destination -Force

