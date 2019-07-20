<#
.SYNOPSIS
Create the first page of the PowerShell Book

.PARAMETER OutputPdfDocument
The path to the output file

.PARAMETER Modulename
The name of the PowerShell module

.NOTES
(c) Andreas Nick Under the MIT License for the Module
https://www.software-virtualisierung.de
https://www.andreasnick.com

We use iTextSharp as library to generate the pdf documents. 
This is licensed under the GNU Affero General Public License.
https://www.nuget.org/packages/iTextSharp/5.5.13.1 
#>
function New-PoShBookTitlePage {

    [OutputType([int])]
    param(
        [Parameter(Mandatory = $true)][System.IO.FileInfo] $OutputPdfDocument,
        [String] $Modulename = "Custom Module",
        [bool] $DisableCredits = $false
        
    )

    if (test-path "$OutputPdfDocument") { Remove-Item "$OutputPdfDocument"  }
    [iTextSharp.text.Document] $Document = New-PDFDocument -File  "$OutputPdfDocument"  -TopMargin $TopMargin -BottomMargin $BottomMargin -LeftMargin $LeftMargin -RightMargin $RightMargin -Author 'The PowerShell Ebook Generator' 


    $ASCII2 = @'
  >>>>>>>                      
     >:::::>                   
        >:::::>                
           >:::::>             
              >:::::>          
                 >:::::>       
                    >:::::>    
                 >:::::>       
              >:::::>          
            >:::::>            
        >:::::>                
     >:::::>     >:::::::::::::::>
  >>>>>>>     >:::::::::::::::>
'@


    $result = $Document.Open() 


    Add-Title -Document $Document -Text "The Powershell Book Generator" -Color 'black' -Centered -FontName "Century Gothic" -FontSize 26

    1..2 | ForEach-Object { Add-NewLine -Document $Document }

    <#
    $p = New-Object -TypeName iTextSharp.text.Paragraph 
    $p.Font = [iTextSharp.text.FontFactory]::GetFont("Courier", $ParagraphFontSize, [iTextSharp.text.Font]::BOLD, [iTextSharp.text.BaseColor]::Black)
    
    $p.SpacingBefore = 2
    $p.SpacingAfter = 2
    $p.Alignment = [iTextSharp.text.Element]::ALIGN_CENTER
    $result = $p.Add($ASCII3) 
    $result = $Document.Add($p) 
    Add-NewLine -Document $Document
    Add-NewLine -Document $Document
   #>

   $pic = [iTextSharp.text.Image]::GetInstance("$PSScriptRoot\PoshLogo.png", [System.Drawing.Imaging.ImageFormat]::Png)
   #$pic.Border = [iTextSharp.text.Rectangle]::BOX
   #$pic.BorderColor = [iTextSharp.text.BaseColor]::BLACK
   #$pic.BorderWidth = 3
   $pic.Alignment = [iTextSharp.text.Image]::ALIGN_CENTER

   
   $null=$document.Add($pic)

    Add-Title -Document $Document -Text "PowerShell with" -Color 'black' -Centered -FontName "Century Gothic" -FontSize 26
    Add-Title -Document $Document -Text "$Modulename" -Color 'black' -Centered -FontName "Century Gothic" -FontSize 26

    if (-not $DisableCredits) {
      1..10 | ForEach-Object { Add-NewLine -Document $Document }

        Add-Text -Document $Document -Text "The PowerShell Book Genrator - Andreas Nick 2019"
        Add-Text -Document $Document -Text "http://www.andreasnick.com"
        Add-Text -Document $Document -Text "http://www.software-virtualisierung.de"
    }

    $Document.Close()
}