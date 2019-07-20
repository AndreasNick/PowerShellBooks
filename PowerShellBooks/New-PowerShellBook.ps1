<#
.SYNOPSIS
Create a PDF book from the PowerSHell documentation of a module

.DESCRIPTION
We use Get-Command [MODULNAME] to get all cmdlets of a PowerShell 
module. Then we generate a documentation with index, title page and page numbering for the module.

.PARAMETER OutputPdfDocument
The path to the output file

.PARAMETER Module
The name of the PowerShell Module

.PARAMETER DisableCredits
If True, the messages about the author on the title page of the PDF document are switched off.

.EXAMPLE
New-PowerShellBook -OutputPdfDocument "c:\temp\PowerShell_With_AppV" -Module "AppVClient"

.NOTES
(c) Andreas Nick Under the MIT License for the Module
https://www.software-virtualisierung.de
https://www.andreasnick.com

We use iTextSharp as library to generate the pdf documents. 
This is licensed under the GNU Affero General Public License.
https://www.nuget.org/packages/iTextSharp/5.5.13.1 

#>

function New-PowerShellBook {

  param(
    [Parameter(Mandatory=$true)][System.IO.FileInfo] $OutputPdfDocument,
    [string] $Module = "NetAdapter",
    [Bool] $DisableCredits = $false
  )

  process {
    $Temp = $env:TEMP
    $TempFolder = "$Temp\PowerShellBooksTemp"
  
    if(test-Path $TempFolder){
      remove-Item -Path $TempFolder -Recurse 
    }
    
    if(-not( test-Path $TempFolder)){
      New-Item -Path $TempFolder -ItemType Directory | Out-Null
    }
  
    $commandlets = Get-Command -Module $Module | Sort-Object
    $Indexliste = @()
    $Currentpage=0
  
    $filelist = New-Object System.Collections.ArrayList
    $filelist.Add($TempFolder+'\title.pdf')  | Out-Null
    $filelist.Add($TempFolder+'\toc.pdf') | Out-Null

    $i=1
    foreach($command in $commandlets){
    
      $OutFile = [String]'{0}\Powershell_CmdLet_{1}.pdf' -f $TempFolder, $Command
      $Pages =  New-PowerShellCommandDocumentation -OutputPdfDocument $OutFile -Command $Command 
  
      $Indexliste += [pscustomobject] @{"Commandlet" = "$Command" ; "Pages" =  $Currentpage }
      $Currentpage += $Pages
      $filelist.Add($OutFile) | Out-Null

      Write-Progress -Activity "Crteate Chapter pdf" -status "Create file $OutFile"  -percentComplete (($i / ($commandlets.count))*100) 
      $i++
    }
    
    $Pages = New-TableOfContent -OutputPdfDocument $($TempFolder+'\toc.pdf') -TOC $Indexliste -AddToPageNumber 3
    
    if($pages -ne 1){
      New-TableOfContent -OutputPdfDocument $($TempFolder+'\toc.pdf') -TOC $Indexliste -AddToPageNumber (2 + $pages)
    }

    New-PoShBookTitlePage -OutputPdfDocument  $($TempFolder+'\title.pdf') -Modulename $Module -DisableCredits $DisableCredits
    New-CombineMultiplePDFs -fileNames $filelist -OutputPdfDocument $($TempFolder+'\combined.pdf')  

    Add-PageNumbers -InputPdfDocument $($TempFolder+'\combined.pdf')  -OutputPdfDocument $OutputPdfDocument

    #Cleanup Temp Folder    
    if(test-Path $TempFolder){
      remove-Item -Path $TempFolder -Recurse 
    }
  }  
}