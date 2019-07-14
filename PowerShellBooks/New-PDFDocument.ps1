Function New-PDFDocument
{
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory=$true)][string] $File,
    [iTextSharp.text.PageSize] $PageSize,
    [int]$TopMargin,
    [int]$BottomMargin,
    [int]$LeftMargin,
    [int]$RightMargin,
    [string]$Author
  )
	
  [iTextSharp.text.Document] $Document = New-Object -TypeName iTextSharp.text.Document 
  
  $null = $Document.SetPageSize([iTextSharp.text.PageSize]::A4)
  $null = $Document.SetMargins($LeftMargin, $RightMargin, $TopMargin, $BottomMargin)
  $global:writer = [iTextSharp.text.pdf.PdfWriter]::GetInstance($Document, [IO.File]::Create($File))
  $null = $Document.AddAuthor($Author) 
  
  Return $Document
  
  
}