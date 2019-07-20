
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
      $Pages =  New-CommandDocumentation -OutputPdfDocument $OutFile -Command $Command 
      #Write-Output $Pages

      $Indexliste += [pscustomobject] @{"Commandlet" = "$Command" ; "Pages" =  $Currentpage }
      $Currentpage += $Pages
      $filelist.Add($OutFile) | Out-Null

      Write-Progress -Activity "Crteate Chapter pdf" -status "Create file $OutFile"  -percentComplete (($i / ($commandlets.count))*100) 
      $i++
    }
    
    $Pages = New-TableOfContent -OutputPdfDocument $($TempFolder+'\toc.pdf') -TOC $Indexliste -AddToPageNumber 3
    #Write-Host $Pages -ForegroundColor Blue
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