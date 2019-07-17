
function New-PoShBookTitlePage {

[OutputType([int])]
param(
  [Parameter(Mandatory=$true)][System.IO.FileInfo] $OutputPdfDocument,
  [String] $Modulename = "Custom Module"
)

if (test-path "$OutputPdfDocument") { Remove-Item "$OutputPdfDocument" -Force }
[iTextSharp.text.Document] $Document = New-PDFDocument -File  "$OutputPdfDocument"  -TopMargin $TopMargin -BottomMargin $BottomMargin -LeftMargin $LeftMargin -RightMargin $RightMargin -Author 'The PowerShell Ebook Generator' 


$ASCII=@'
        _                                   
       (_) _                                
          (_) _                             
             (_) _                          
              _ (_)                         
           _ (_)                            
        _ (_)_  _  _  _ 
       (_)  (_)(_)(_)(_)

'@

$ASCII2=@'
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
     >:::::>  >:::::::::::::::>                              
  >>>>>>>     >>>>>>>>>>>>>>>>>  
'@

$result = $Document.Open() 


1..10 | ForEach-Object {Add-NewLine -Document $Document }

$p = New-Object -TypeName iTextSharp.text.Paragraph 
$p.Font = [iTextSharp.text.FontFactory]::GetFont("Courier", $ParagraphFontSize, [iTextSharp.text.Font]::BOLD, [iTextSharp.text.BaseColor]::Black)
$p.SpacingBefore = 2
$p.SpacingAfter = 2
$result = $p.Add($Text) 
$result = $Document.Add($p) 


Add-Title -Document $Document -Text "Powershell E-Book" -Color 'black' -Centered -FontName "Century Gothic" -FontSize 26
Add-Title -Document $Document -Text "Powershell Module: $Modulename " -Color 'black' -Centered -FontName "Century Gothic" -FontSize 26

$Document.Close()
}