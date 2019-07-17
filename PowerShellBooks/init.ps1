
# use this file to define global variables on module scope
# or perform other initialization procedures.
# this file will not be touched when new functions are exported to
# this module.

#requires -Version 3

$Global:ParagraphFontSize = 11
$Global:HeadlineFontSize = 18
$Global:SecondHeadlineFontSize = 16
$Global:LeftMargin = 80
$Global:RightMargin = 80

$Global:TopMargin = 80
$Global:BottomMargin = 80

Unblock-File -Path "$PSScriptRoot\lib\itextsharp.dll" -Confirm:$false
Add-Type -Path "$PSScriptRoot\lib\itextsharp.dll"

