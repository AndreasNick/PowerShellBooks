
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
    [PSCustomObject[]] $TOC,
    [int] $AddToPageNumber = 0 #Titlepage + Toc = 2
    
  )

  if (test-path "$OutputPdfDocument") { Remove-Item "$OutputPdfDocument"  }
  [iTextSharp.text.Document] $Document = New-PDFDocument -File  "$OutputPdfDocument"  -TopMargin $TopMargin -BottomMargin $BottomMargin -LeftMargin $LeftMargin -RightMargin $RightMargin -Author 'The PowerShell Ebook Generator' 

  $result = $Document.Open() 
  
  $Null=Add-Headline -Document $Document -Text $("Table Of Contents")
  $Null=Add-NewLine -Document $Document

  $dottedLine = new-object  iTextSharp.text.Chunk  ( new-object iTextSharp.text.pdf.draw.DottedLineSeparator );
  #[iTextSharp.text.Paragraph] $p

  foreach($entry in $TOC){
    $p = New-Object iTextSharp.text.Paragraph -ArgumentList  @($entry.Commandlet)
    $null=$p.Add($dottedLine)
    $null=$p.Add($entry.pages + $AddToPageNumber)
    $null=$Document.add($p);
  }

  $Pages =  $global:writer.CurrentPageNumber

  $Document.close();
  return $Pages
}