<#
.SYNOPSIS
This command binds several pdf documents together

.DESCRIPTION
This command binds several pdf documents together
I included this command in the module, because otherwise 
I didn't find anything for PowerShell. It will surely help one or the other.

.PARAMETER fileNames
A List of PDF Files zu combine

.PARAMETER OutputPdfDocument
The path to the output file

.EXAMPLE
New-CombineMultiplePDFs -fileNames @('c:\temp\file1.pdf','c:\temp\file2.pdf') -OutputPdfDocument 'c:\temp\combined.pdf' 

.NOTES
(c) Andreas Nick Under the MIT License for the Module
https://www.software-virtualisierung.de
https://www.andreasnick.com

We use iTextSharp as library to generate the pdf documents. 
This is licensed under the GNU Affero General Public License.
https://www.nuget.org/packages/iTextSharp/5.5.13.1 

#>

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

