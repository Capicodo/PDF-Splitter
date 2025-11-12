$source = ".\dist\Monatsbericht Automat Test.exe"
$destination = "\\hokkaido\Daten\IT\Mu\Tools\Monatsbericht Automat\Monatsbericht Automat Test.exe"

Copy-Item $source -Destination $destination -Force
