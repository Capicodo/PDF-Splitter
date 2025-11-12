$source = ".\dist\Monatsbericht Automat.exe"
$destination = "\\hokkaido\Daten\IT\Mu\Tools\Monatsbericht Automat\Monatsbericht Automat.exe"

Copy-Item $source -Destination $destination -Force
