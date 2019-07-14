function Add-Headline{
  [CmdletBinding()]
  param(
    [iTextSharp.text.Document]$Document, 
    [string] $Text
  )
  
  Add-Text -Document $Document -FontName "Arial" -FontSize $HeadlineFontSize -Text $Text
  
}