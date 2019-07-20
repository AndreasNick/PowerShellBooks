<#
.SYNOPSIS
Create a Text in a PDF document

.PARAMETER Document
PDF Document

.PARAMETER Text
The text

.PARAMETER Centered
Titel shut be centered?

.PARAMETER FontName
Name of the font (Arial, etc.)

.PARAMETER FontSize
Maybe 8

.PARAMETER Color
[iTextSharp.text.BaseColor]::Black

.NOTES
(c) Andreas Nick Under the MIT License for the Module
https://www.software-virtualisierung.de
https://www.andreasnick.com

We use iTextSharp as library to generate the pdf documents. 
This is licensed under the GNU Affero General Public License.
https://www.nuget.org/packages/iTextSharp/5.5.13.1 
#>
function Add-Text
{
  [CmdletBinding()]
  param
  (
    [iTextSharp.text.Document]$Document,
    [string]$Text,
    [string]$FontName = 'Arial',
    [int]$FontSize = 11,
    [string]$Color = 'BLACK'
  )
	
  $p = New-Object -TypeName iTextSharp.text.Paragraph 
  $p.Font = [iTextSharp.text.FontFactory]::GetFont($FontName, $FontSize, [iTextSharp.text.Font]::NORMAL, [iTextSharp.text.BaseColor]::$Color) 
  
  $p.SpacingBefore = 2
  $p.SpacingAfter = 2
  
  $result = $p.Add($Text) 
  $result = $Document.Add($p) 
  
  
}