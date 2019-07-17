
function New-PowerShellBook {
  

  param(
    [Parameter(Mandatory=$true)][System.IO.FileInfo] $OutputPdfDocument,
    [string] $Module = "NetAdapter"
  )

  process {
    $Temp = $env:TEMP
    $TempFolder = "$Temp\PowerShellBooksTemp"
  
    
    if(-not( test-Path $TempFolder)){
      New-Item -Path $TempFolder -ItemType Directory
    }
  
    
    $commandlets = Get-Command -Module $Module | Sort-Object
    
    
    $Indexliste = @()
    #$Currentpage=0
  
    foreach($command in $commandlets){
    
      $OutFile = '{0}\Powershell_With_{1}.pdf' -f $TempFolder, $Command
      $Pages =  New-CommandDocumentation -OutputPdfDocument $OutFile -Command $Command 
      #Write-Output $Pages

      $Indexliste += [pscustomobject] @{"Commandlet" = "$Command" ; "Pages" =  $Pages }
      #Write-Output $($Pages.count.ToString() + " ->  $pages")
    }
    
    if(test-Path $TempFolder){
      remove-Item -Path $TempFolder -Recurse 
    }
    
    New-TableOfContent -OutputPdfDocument C:\Users\Andreas\Desktop\toc.pdf -TOC $Indexliste

    $Indexliste
  }  
    
}