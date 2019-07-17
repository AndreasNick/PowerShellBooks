function Add-NewLine
{
  [CmdletBinding()]
  param ([iTextSharp.text.Document] $Document)
  $result = $Document.Add( (New-Object iTextSharp.text.Paragraph "`n" )) | Out-Null
  
}