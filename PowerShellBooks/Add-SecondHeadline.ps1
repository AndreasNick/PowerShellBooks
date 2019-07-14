function Add-SecondHeadline{
  [CmdletBinding()]
  param(
    [iTextSharp.text.Document]$Document, 
    [string] $Text
  )
  Add-Text -Document $Document -FontName "Arial" -FontSize $SecondHeadlineFontSize -Text $Text
  }