# The PowerShell Book Generator
After seeing how some colleagues use the built-in PowerShell help to publish and sell whole books, I thought you'd have to build a PDF genrator for it. And here is this generator. The PowerShell module PowerShellBooks. We use Get-Command MODULNAME to get all cmdlets of a PowerShell module. Then we generate a documentation with index, title page and page numbering for the module. The whole package is of course also suitable to document your PowerShell modules. The credits can of course be hidden.

## What else is important
What else is important. If you want it to go fast, only for the module you want to document
```powershell
Update-Help -Module Storage
```
or for all
```powershell
Update-Help -Module * -force
```

## How does it work?

First the module must be imported. This can also be in the module memory, for example. I am also planning to store a version in the PowerShell gallery.

```powershell
#In the Script folder
Import-Module $PSScriptRoot\PowerShellBooks -Force
Import-Module 'YOURPATH\'+'\PowerShellBooks') 

#In the module path (PowerShell Gallerie)
Install-Module 'PowerShellBooks'
```

The CmdLet is then called.

```powershell
#For Example
New-PowerShellBook -OutputPdfDocument "c:\temp\PowerShell_With_AppV" -Module "AppVClient"
```
The whole thing also works for single cmdlets, of course.

```powershell
#For Example
New-CommandDocumentation -OutputPdfDocument "c:\temp\get-ChildItems_Doc.pdf -Command "Get-ChildItems" 
```
Since the function is so practical I also released the CmdLet to combine PDF documents.
```powershell
#For Example
New-CombineMultiplePDFs -fileNames @('c:\temp\file1.pdf','c:\temp\file2.pdf') -OutputPdfDocument 'c:\temp\combined.pdf' 
```
For colleagues who prefer to use a GUI there is also a very simple one included

```powershell
PowerShellBooksGui.ps1 
```

![GUI](https://github.com/AndreasNick/PowerShellBooks/blob/master/ThePowerShellBookGeneratorGUI.jpg?raw=true)
       