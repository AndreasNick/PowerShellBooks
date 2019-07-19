function New-CombineMultiplePDFs
{
    param([string[]] $fileNames, [System.IO.FileInfo] $OutputPdfDocument)

    $document = New-Object  Document()
    if (test-path "$OutputPdfDocument") { Remove-Item "$OutputPdfDocument"  }
    [iTextSharp.text.Document] $Document = New-PDFDocument -File  "$OutputPdfDocument"  -Author 'The PowerShell Ebook Generator' 
    
    $document.Open()

       foreach ($fileName in $fileNames)
       {
           [iTextSharp.text.pdf.PdfReader] $reader = iTextSharp.text.pdf.PdfReader $fileName
           $reader.ConsolidateNamedDestinations();

           for ($i = 1; $i -le $reader.NumberOfPages; $i++)
           {                
               $page = $global:writer.GetImportedPage($reader, $i)
               $global:writer.AddPage($page)
           }
           
           <#
           PRAcroForm form = reader.AcroForm;
           if (form != null)
           {
               writer.CopyAcroForm(reader);
           }
           #>
           $reader.Close()
       }

       $Global:writer.Close()
       $Global:document.Close();
   
}

