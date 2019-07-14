function Add-Title
{
  [CmdletBinding()]
  param
  (
    [iTextSharp.text.Document]$Document,
    [string]$Text,
    [switch]$Centered,
    [string]$FontName = 'Arial',
    [int]$FontSize = 16,
    [string]$Color = 'BLACK'
  )
	
  $p = New-Object -TypeName iTextSharp.text.Paragraph
  
  $p.Font = [iTextSharp.text.FontFactory]::GetFont($FontName, $FontSize, [iTextSharp.text.Font]::BOLD, [iTextSharp.text.BaseColor]::$Color)
  if ($Centered)
  {
    $p.Alignment = [iTextSharp.text.Element]::ALIGN_CENTER
  }
  $p.SpacingBefore = 5
  $p.SpacingAfter = 5
  $result = $p.Add($Text) 
  $result = $Document.Add($p) 
}