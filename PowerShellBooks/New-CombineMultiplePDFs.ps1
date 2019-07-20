function New-CombineMultiplePDFs
{
  param(
    [string[]] $fileNames, 
    [System.IO.FileInfo] $OutputPdfDocument
  )
  
  if (test-path "$OutputPdfDocument") { Remove-Item "$OutputPdfDocument"  }
  
  $fileStream = New-Object System.IO.FileStream($OutputPdfDocument, [System.IO.FileMode]::OpenOrCreate)
  $document = New-Object iTextSharp.text.Document
  $pdfCopy = New-Object iTextSharp.text.pdf.PdfCopy($document, $fileStream)
    
   
  $document.Open()
  
  foreach ($fileName in $fileNames)
  {
    [System.IO.FileInfo] $fi = $fileName
    $reader = New-Object iTextSharp.text.pdf.PdfReader -argumentlist $fi.fullname
    
    $pdfCopy.AddDocument($reader);

    #[iTextSharp.text.pdf.PdfReader] $reader = iTextSharp.text.pdf.PdfReader $fileName
    #$pdfcopy = New-Object iTextSharp.text.pdf.PdfCopy -ArgumentList $fileName
           
    <#
        PRAcroForm form = reader.AcroForm;
        if (form != null)
        {
        writer.CopyAcroForm(reader);
        }
    #>
    $reader.Close()
  }

  $pdfCopy.Close();
  $document.Close();
  $fileStream.Close(); 
}

