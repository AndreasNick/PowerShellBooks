
function New-TableOfContent {
<#
.SYNOPSIS
Create a Tabel Of Content for as list with pscustomobjects

.DESCRIPTION
Create a Tabel Of Content for as list with pscustomobjects

.PARAMETER OutputPdfDocument
The path to the output file

.PARAMETER TOC
[PSCustomObject[]] $TOC a List with entries @{Commandlet=''; Pages=''} for the toc

.EXAMPLE
New-TableOfContent -OutputPdfDocument $($TempFolder+'\toc.pdf') -TOC $Indexliste -AddToPageNumber (2 + $pages)

.NOTES
(c) Andreas Nick Under the MIT License for the Module
https://www.software-virtualisierung.de
https://www.andreasnick.com

We use iTextSharp as library to generate the pdf documents. 
This is licensed under the GNU Affero General Public License.
https://www.nuget.org/packages/iTextSharp/5.5.13.1 

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