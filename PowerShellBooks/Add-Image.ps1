function Add-Image
{
  [CmdletBinding()]
  param
  (
    [iTextSharp.text.Document] $Document,
    [string]$File,
    [int]$Scale = 100
  )
	
  [iTextSharp.text.Image]$img = [iTextSharp.text.Image]::GetInstance($File)
  $img.ScalePercent(50)
  $Document.Add($img)
}