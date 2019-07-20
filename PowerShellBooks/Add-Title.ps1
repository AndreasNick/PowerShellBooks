<#
.SYNOPSIS
Create a Title in a PDF document

.PARAMETER Document
PDF Document

.PARAMETER Text
The text of the title

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

function Add-Title
{
  [CmdletBinding()]
  param
  (
    [iTextSharp.text.Document]$Document,
    [string]$Text,
    [switch]$Centered,
    [string]$FontName = 'Arial',
    [int]$FontSize = 16,
    [string]$Color = 'BLACK'
  )
	
  $p = New-Object -TypeName iTextSharp.text.Paragraph
  
  $p.Font = [iTextSharp.text.FontFactory]::GetFont($FontName, $FontSize, [iTextSharp.text.Font]::BOLD, [iTextSharp.text.BaseColor]::$Color)
  if ($Centered)
  {
    $p.Alignment = [iTextSharp.text.Element]::ALIGN_CENTER
  }
  $p.SpacingBefore = 5
  $p.SpacingAfter = 5
  $result = $p.Add($Text) 
  $result = $Document.Add($p) 
}