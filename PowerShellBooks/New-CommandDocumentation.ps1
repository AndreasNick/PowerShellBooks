
function New-CommandDocumentation {
    <#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER OutputPdfDocument
Parameter description

.PARAMETER Command
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>

    [CmdletBinding()]
    param(
        [System.IO.FileInfo] $OutputPdfDocument,
        [string] $Command = "Get-Location"
    )

    if (test-path "$OutputPdfDocument") { Remove-Item "$OutputPdfDocument" -Force }
    [iTextSharp.text.Document] $Document = New-PDFDocument -File  "$OutputPdfDocument"  -TopMargin 20 -BottomMargin 20 -LeftMargin $LeftMargin -RightMargin $RightMargin -Author 'The PowerShell Ebook Generator' 

    $result = $Document.Open() 

  
    $helpText = Get-Help $command -Full

    Add-Headline -Document $Document -Text $("Cmdlet: " + $helpText.Name)
    Add-NewLine -Document $Document
    Add-SecondHeadline -Document $Document -Text "Synops"
    Add-Text -Document $Document -Text $helpText.Synopsis
  
    Add-NewLine -Document $Document

    Add-SecondHeadline -Document $Document -Text "Syntax"
  
    $syntax = $($helpText.syntax | Out-String)
    $result = ""
  
  
    foreach ($line in @($syntax.Split("`n"))) {
        if ($line.length -gt 2) {
            $result += $line
        }
    }
  
  
    Add-text -Document $Document $result -FontName "Courier"
  
    Add-NewLine -Document $Document
  
    Add-SecondHeadline -Document $Document -Text "Description"
  
    foreach ($desc in $helpText.description) {
        Add-Text -Document $Document -Text $desc.Text
    }

    Add-NewLine -Document $Document 
  
    Add-SecondHeadline -Document $Document -Text "Parameters"
    
  
    # Build Parameter Table
    #Required?                    false        
    #Position?                    named        
    #Default value                none        
    #Accept pipeline input?       false        
    #Accept wildcard characters?  false 
  
    #$helpText.parameters.parameter[5] | fl
    
    foreach ($para in $helpText.parameters.parameter) {
  
        $p = New-Object -TypeName iTextSharp.text.Paragraph
        $Font = [iTextSharp.text.FontFactory]::GetFont("Arial", $ParagraphFontSize, [iTextSharp.text.Font]::NORMAL, [iTextSharp.text.BaseColor]::BLACK)
        $FatFont = [iTextSharp.text.FontFactory]::GetFont("Arial", $ParagraphFontSize, [iTextSharp.text.Font]::BOLD, [iTextSharp.text.BaseColor]::Black)
    
        $p.SpacingBefore = 2
        $p.SpacingAfter = 2
        $result = $p.Add((New-Object iTextSharp.text.Chunk  $("Parameter :"), $font))
        $result = $p.Add((New-Object iTextSharp.text.Chunk  $($para.name), $FatFont))
        $result = $Document.Add($p)
        Add-Text -Document $Document -Text $("Description :" + $para.description[0].Text)
        Add-NewLine -Document $Document
        $parTable = New-Object System.Collections.ArrayList  
        
        $result = $parTable.add("Required")
        $result = $parTable.add($para.required)
        $result = $parTable.add("Position")
        $result = $parTable.add($para.position)
        $result = $parTable.add("Default value")
        $result = $parTable.add($para.DefaultValue)
        $result = $parTable.add('Accept pipeline input?')
        $result = $parTable.add($para.PipelineInput)  
        $result = $parTable.add('Accept wildcard characters?')
        $result = $parTable.add($para.globbing)         

        $result = Add-Table -Document $Document -Dataset $parTable -Cols 2 -Centered
        Add-NewLine  -Document $Document
    
    }
   
   
    if ($helpText.inputTypes -ne $null ) {
        if ($helpText.inputTypes.inputType.description.Length -ne 0) {
            Add-SecondHeadline -Document $Document "Inputs"
            foreach ($desc in $helpText.inputTypes.inputType.description) {
        
                Add-Text -Document $Document -text $desc.text
            }
            Add-Text -Document $Document -text $("Type : " + $helpText.inputTypes.inputType.type.name)
            Add-NewLine -Document $Document
        }
    
    }

    if ($helpText.returnValues -ne $null ) {
        if ($helpText.returnValues.returnValue.description.Length -ne 0) {
            Add-SecondHeadline -Document $Document "Outputs"
            foreach ($desc in $helpText.returnValues.returnValue.description) {
        
                Add-Text -Document $Document -text $desc.text
            }
            Add-Text -Document $Document -text $("Type : " + $helpText.returnValues.returnValue.type.name)
            Add-NewLine -Document $Document
        }
    
    }
  
   
    if ($helpText.alertSet.alert.Count -ne 0) {

        Add-SecondHeadline -Document $Document "Notes" 
    
        foreach ($note in $helpText.alertSet.alert) {
            Add-Text -Document $Document $($note.Text).Trim() 
            Add-NewLine -Document $Document 
        } 

    }
  
  
    if ($helpText.examples.example.Count -gt 0) {
        Add-SecondHeadline -Document $Document "Examples"
        foreach ($example in $helpText.examples.example) {
            Add-Text -Document $Document $example.title
            foreach ($remarks in $example.Remarks) {
                Add-Text -Document $Document -Text $remarks.Text 
            }
      
            #Code
            $result = Add-Table -Document $Document -Dataset @($example.Code + "`n") -UsegrayBG -Cols 1 -Centered  -UseConsoleFont -Noborder -WidthPercentage 95
            Add-NewLine  -Document $Document
      
            Add-NewLine -Document $Document
    
        }
    
    }
    
    #$pdf.PageNumber
    #$global:writer.CurrentPageNumber

    Add-NewLine -Document $Document
    $result = $Document.Close() 

}