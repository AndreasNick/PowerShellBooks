<#

Test script for the PowerShell Book Generator.

The PowerShell Book Generator creates an A4 PDF document with full help on a PowerShell module. There is also a table of contents and a title page.

(c) Andreas Nick 2019

#>

Import-Module $PSScriptRoot\PowerShellBooks -Force

Import-Module 'C:\Users\Andreas\Desktop\Signal-IdunaCitrixPublishing\Signal-IdunaCitrixPublishing\CitrixPoshHelper' -Force

#$Module = 'Storage'
$Module = 'CitrixPoshHelper'

$OutFile = '{0}\desktop\Powershell_With_{1}.pdf' -f $env:userprofile, $Module

New-PowerShellBook -OutputPdfDocument $OutFile -Module $Module


