Import-Module $PSScriptRoot\PowerShellBooks -Force



$Command = 'DISM'

$OutFile = '{0}\desktop\Powershell_With_{1}.pdf' -f $env:userprofile, $Command

New-CommandDocumentation -OutputPdfDocument $OutFile -Command $Command 



#$pdf.AddHeader((New-Object iTextSharp.text.Paragraph "Header"))
<#

    #Add-Image -Document $pdf -File "$PSScriptRoot\kitten1.jpg"
    Add-Table -Document $pdf -Dataset @('Name', 'Mittens', 'Age', '1.5', 'Fur color', 'Black and white', 'Favorite toy', 'String') -Cols 2 -Centered
    #Add-Image -Document $pdf -File "$PSScriptRoot\kitten2.jpg"
    Add-Table -Document $pdf -Dataset @('Name', 'Achilles', 'Age', '2', 'Fur color', 'Grey', 'Favorite toy', 'Shoes') -Cols 2 -Centered
    Add-Text -Document $pdf -Text 'Meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow meow'
#>


