function Add-Table
{
  [CmdletBinding()]
  param
  (
    [iTextSharp.text.Document]$Document,
    [String[]]$Dataset,
    [int]$Cols = 3,
    [switch]$Centered,
    [switch] $UsegrayBG = $false,
    [switch] $UseConsoleFont = $false,
    [switch] $Noborder = $false,
    [int] $WidthPercentage = 0
    
  )
	
  $t = New-Object -TypeName iTextSharp.text.pdf.PDFPTable -ArgumentList ($Cols)
  if($WidthPercentage -ne 0){
    $t.WidthPercentage = $WidthPercentage
  }
  $Gray = new-object iTextSharp.text.BaseColor 240, 240, 240
  $ConsoleFont = [iTextSharp.text.FontFactory]::GetFont("Courier", $ParagraphFontSize, [iTextSharp.text.Font]::BOLD, [iTextSharp.text.BaseColor]::BLACK)
  
  $t.SpacingBefore = 5
  $t.SpacingAfter = 5
  if (!$Centered)
  {
    $t.HorizontalAlignment = 0
  }
  foreach ($data in $Dataset)
  {
    $p = $null 
    
    if($UseConsoleFont){
      $p = New-Object -TypeName iTextSharp.text.Phrase  $data, $ConsoleFont
    } else {
      $p = New-Object -TypeName iTextSharp.text.Phrase  $data
      
    }
    
    #$t.AddCell($data)
    $cell = New-Object iTextSharp.text.pdf.PdfPCell $p
    
    if($UsegrayBG){
      $cell.BackgroundColor = $Gray
    }

    if($Noborder){
      $cell.Border = $null
    }
    
    $t.AddCell($cell)
  }
  $Document.Add($t)
}