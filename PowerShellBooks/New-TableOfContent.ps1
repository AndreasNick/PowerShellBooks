
function New-TableOfContent {
<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER OutputPdfDocument
Parameter description

.PARAMETER TOC
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>

  [OutputType([int])]
  param(
    [Parameter(Mandatory=$true)][System.IO.FileInfo] $OutputPdfDocument,
    [PSCustomObject[]] $TOC
    
  )

  if (test-path "$OutputPdfDocument") { Remove-Item "$OutputPdfDocument" -Force }
  [iTextSharp.text.Document] $Document = New-PDFDocument -File  "$OutputPdfDocument"  -TopMargin 20 -BottomMargin 20 -LeftMargin $LeftMargin -RightMargin $RightMargin -Author 'The PowerShell Ebook Generator' 

  $result = $Document.Open() 
  
  $dottedLine = new-object  iTextSharp.text.Chunk  (new-object iTextSharp.text.pdf.draw.DottedLineSeparator());
  [iTextSharp.text.Paragraph] $p

  foreach($entry in $TOC){
    $p = New-Object iTextSharp.text.Paragraph -ArgumentList  @($entry.Commandlet)
    $p.Add($dottedLine)
    $p.Add($entry.pages)
    $Document.add($p);
  }

  $Document.close();

}