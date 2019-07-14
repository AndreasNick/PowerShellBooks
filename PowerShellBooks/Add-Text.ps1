function Add-Text
{
  [CmdletBinding()]
  param
  (
    [iTextSharp.text.Document]$Document,
    [string]$Text,
    [string]$FontName = 'Arial',
    [int]$FontSize = 11,
    [string]$Color = 'BLACK'
  )
	
  $p = New-Object -TypeName iTextSharp.text.Paragraph 
  $p.Font = [iTextSharp.text.FontFactory]::GetFont($FontName, $FontSize, [iTextSharp.text.Font]::NORMAL, [iTextSharp.text.BaseColor]::$Color) 
  
  $p.SpacingBefore = 2
  $p.SpacingAfter = 2
  
  $result = $p.Add($Text) 
  $result = $Document.Add($p) 
  
  
}