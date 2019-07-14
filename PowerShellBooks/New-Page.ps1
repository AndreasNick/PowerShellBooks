function New-Page
{
  [CmdletBinding()]
  param ([iTextSharp.text.Document] $Document)
  
  $result=$Document.NewPage() 

}