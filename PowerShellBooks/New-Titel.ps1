function New-Titel{
  [CmdletBinding()]
  param([iTextSharp.text.Document]$Document)

  1..10 | ForEach-Object {Add-NewLine -Document $pdf }

  Add-Title -Document $pdf -Text "Powershell E-Book" -Color 'black' -Centered -FontName "Century Gothic" -FontSize 26
  Add-Title -Document $pdf -Text "Powershell Module: $Module" -Color 'black' -Centered -FontName "Century Gothic" -FontSize 26
  New-Page -Document $pdf


}