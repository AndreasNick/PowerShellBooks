﻿<#
.SYNOPSIS
Short description

.DESCRIPTION
create a pdf document about a PowerShell function or cmdlet

.PARAMETER OutputPdfDocument
The path to the output file

.PARAMETER Command
The order. For example, "Get-ChildItem"

.EXAMPLE
#The command return the Number of Pages
$pages = New-CommandDocumentation -OutputPdfDocument "c:\temp\get-ChildItems_Doc.pdf -Command "Get-ChildItems" 

.NOTES
(c) Andreas Nick Under the MIT License for the Module
https://www.software-virtualisierung.de
https://www.andreasnick.com

We use iTextSharp as library to generate the pdf documents. 
This is licensed under the GNU Affero General Public License.
https://www.nuget.org/packages/iTextSharp/5.5.13.1 

#>

function New-PowerShellCommandDocumentation {
  [OutputType([int])]
  param(
    [Parameter(Mandatory=$true)][System.IO.FileInfo] $OutputPdfDocument,
    [string] $Command = "Get-Location"
  )
    if (test-path "$OutputPdfDocument") { Remove-Item "$OutputPdfDocument"  }
    [iTextSharp.text.Document] $Document = New-PDFDocument -File  "$OutputPdfDocument"  -TopMargin $TopMargin -BottomMargin $BottomMargin -LeftMargin $LeftMargin -RightMargin $RightMargin -Author 'The PowerShell Ebook Generator' 

    $result = $Document.Open() 
    
    $helpText = Get-Help $command -Full

    Add-Headline -Document $Document -Text $("Cmdlet: " + $helpText.Name)

    Add-NewLine -Document $Document
    Add-SecondHeadline -Document $Document -Text "Synops"
    Add-Text -Document $Document -Text $helpText.Synopsis
  
    Add-NewLine -Document $Document

    Add-SecondHeadline -Document $Document -Text "Syntax"
  
    #$syntax = $($helpText.syntax | Out-String)
    $syntax =  $($helpText | Out-String) #-replace "`n","" -replace "`r",""
    $pattern = '(?ms)SYNTAX(.+?)(.|\s)(DESCRIPTION|PARAMETERS|ALIASE|BESCHREIBUNG)'
    $syntax = [regex]::Match($syntax, $pattern).Groups[1].Value

   
    foreach ($line in @($syntax.Split("`n"))) {
      if ($line.length -gt 2) {
        $result += $line
      }
    }
   
    $result = $result -replace $Command, $("`n" + $Command)
    $result = $result.substring(2)
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
        if ($para.description.count -ge 1 ) {
            Add-Text -Document $Document -Text $("Description :" + $para.description[0].Text)
            Add-NewLine -Document $Document
        }

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
        if($node -ne $null){
        Add-Text -Document $Document $($note.Text).Trim() 
        Add-NewLine -Document $Document 
        }
      } 

    }
  
  
    if ($null -ne $helpText.examples.example) {
      Add-SecondHeadline -Document $Document "Examples"
      foreach ($example in $helpText.examples.example) {
        $title =  $($example.title) -replace 'BEISPIEL', 'EXAMPLE'
        
        Add-Text -Document $Document  $title 
        foreach ($remarks in $example.Remarks) {
          Add-Text -Document $Document -Text $remarks.Text 
        }
      
        #Code
        $result = Add-Table -Document $Document -Dataset @($example.Code + "`n") -UsegrayBG -Cols 1 -Centered  -UseConsoleFont -Noborder -WidthPercentage 95
        Add-NewLine -Document $Document
        Add-NewLine -Document $Document
      }
    }
    
    $Pages =  $global:writer.CurrentPageNumber
 
    Add-NewLine -Document $Document
    $result = $Document.Close() 
      return $Pages
    
  #$Pages

}
# SIG # Begin signature block
# MIIe2gYJKoZIhvcNAQcCoIIeyzCCHscCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAYz4eCAo9yzTZt
# lFeH0S2g5pZj9QNrVkuYZvDl3CvxAqCCGb8wggSEMIIDbKADAgECAhBCGvKUCYQZ
# H1IKS8YkJqdLMA0GCSqGSIb3DQEBBQUAMG8xCzAJBgNVBAYTAlNFMRQwEgYDVQQK
# EwtBZGRUcnVzdCBBQjEmMCQGA1UECxMdQWRkVHJ1c3QgRXh0ZXJuYWwgVFRQIE5l
# dHdvcmsxIjAgBgNVBAMTGUFkZFRydXN0IEV4dGVybmFsIENBIFJvb3QwHhcNMDUw
# NjA3MDgwOTEwWhcNMjAwNTMwMTA0ODM4WjCBlTELMAkGA1UEBhMCVVMxCzAJBgNV
# BAgTAlVUMRcwFQYDVQQHEw5TYWx0IExha2UgQ2l0eTEeMBwGA1UEChMVVGhlIFVT
# RVJUUlVTVCBOZXR3b3JrMSEwHwYDVQQLExhodHRwOi8vd3d3LnVzZXJ0cnVzdC5j
# b20xHTAbBgNVBAMTFFVUTi1VU0VSRmlyc3QtT2JqZWN0MIIBIjANBgkqhkiG9w0B
# AQEFAAOCAQ8AMIIBCgKCAQEAzqqBP6OjYXiqMQBVlRGeJw8fHN86m4JoMMBKYR3x
# Lw76vnn3pSPvVVGWhM3b47luPjHYCiBnx/TZv5TrRwQ+As4qol2HBAn2MJ0Yipey
# qhz8QdKhNsv7PZG659lwNfrk55DDm6Ob0zz1Epl3sbcJ4GjmHLjzlGOIamr+C3bJ
# vvQi5Ge5qxped8GFB90NbL/uBsd3akGepw/X++6UF7f8hb6kq8QcMd3XttHk8O/f
# Fo+yUpPXodSJoQcuv+EBEkIeGuHYlTTbZHko/7ouEcLl6FuSSPtHC8Js2q0yg0Hz
# peVBcP1lkG36+lHE+b2WKxkELNNtp9zwf2+DZeJqq4eGdQIDAQABo4H0MIHxMB8G
# A1UdIwQYMBaAFK29mHo0tCb3+sQmVO8DveAky1QaMB0GA1UdDgQWBBTa7WR0FJwU
# PKvdmam9WyhNizzJ2DAOBgNVHQ8BAf8EBAMCAQYwDwYDVR0TAQH/BAUwAwEB/zAR
# BgNVHSAECjAIMAYGBFUdIAAwRAYDVR0fBD0wOzA5oDegNYYzaHR0cDovL2NybC51
# c2VydHJ1c3QuY29tL0FkZFRydXN0RXh0ZXJuYWxDQVJvb3QuY3JsMDUGCCsGAQUF
# BwEBBCkwJzAlBggrBgEFBQcwAYYZaHR0cDovL29jc3AudXNlcnRydXN0LmNvbTAN
# BgkqhkiG9w0BAQUFAAOCAQEATUIvpsGK6weAkFhGjPgZOWYqPFosbc/U2YdVjXkL
# Eoh7QI/Vx/hLjVUWY623V9w7K73TwU8eA4dLRJvj4kBFJvMmSStqhPFUetRC2vzT
# artmfsqe6um73AfHw5JOgzyBSZ+S1TIJ6kkuoRFxmjbSxU5otssOGyUWr2zeXXbY
# H3KxkyaGF9sY3q9F6d/7mK8UGO2kXvaJlEXwVQRK3f8n3QZKQPa0vPHkD5kCu/1d
# Di4owb47Xxo/lxCEvBY+2KOcYx1my1xf2j7zDwoJNSLb28A/APnmDV1n0f2gHgMr
# 2UD3vsyHZlSApqO49Rli1dImsZgm7prLRKdFWoGVFRr1UTCCBOYwggPOoAMCAQIC
# EGJcTZCM1UL7qy6lcz/xVBkwDQYJKoZIhvcNAQEFBQAwgZUxCzAJBgNVBAYTAlVT
# MQswCQYDVQQIEwJVVDEXMBUGA1UEBxMOU2FsdCBMYWtlIENpdHkxHjAcBgNVBAoT
# FVRoZSBVU0VSVFJVU1QgTmV0d29yazEhMB8GA1UECxMYaHR0cDovL3d3dy51c2Vy
# dHJ1c3QuY29tMR0wGwYDVQQDExRVVE4tVVNFUkZpcnN0LU9iamVjdDAeFw0xMTA0
# MjcwMDAwMDBaFw0yMDA1MzAxMDQ4MzhaMHoxCzAJBgNVBAYTAkdCMRswGQYDVQQI
# ExJHcmVhdGVyIE1hbmNoZXN0ZXIxEDAOBgNVBAcTB1NhbGZvcmQxGjAYBgNVBAoT
# EUNPTU9ETyBDQSBMaW1pdGVkMSAwHgYDVQQDExdDT01PRE8gVGltZSBTdGFtcGlu
# ZyBDQTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAKqC8YSpW9hxtdJd
# K+30EyAM+Zvp0Y90Xm7u6ylI2Mi+LOsKYWDMvZKNfN10uwqeaE6qdSRzJ6438xqC
# pW24yAlGTH6hg+niA2CkIRAnQJpZ4W2vPoKvIWlZbWPMzrH2Fpp5g5c6HQyvyX3R
# TtjDRqGlmKpgzlXUEhHzOwtsxoi6lS7voEZFOXys6eOt6FeXX/77wgmN/o6apT9Z
# RvzHLV2Eh/BvWCbD8EL8Vd5lvmc4Y7MRsaEl7ambvkjfTHfAqhkLtv1Kjyx5VbH+
# WVpabVWLHEP2sVVyKYlNQD++f0kBXTybXAj7yuJ1FQWTnQhi/7oN26r4tb8QMspy
# 6ggmzRkCAwEAAaOCAUowggFGMB8GA1UdIwQYMBaAFNrtZHQUnBQ8q92Zqb1bKE2L
# PMnYMB0GA1UdDgQWBBRkIoa2SonJBA/QBFiSK7NuPR4nbDAOBgNVHQ8BAf8EBAMC
# AQYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNVHSUEDDAKBggrBgEFBQcDCDARBgNV
# HSAECjAIMAYGBFUdIAAwQgYDVR0fBDswOTA3oDWgM4YxaHR0cDovL2NybC51c2Vy
# dHJ1c3QuY29tL1VUTi1VU0VSRmlyc3QtT2JqZWN0LmNybDB0BggrBgEFBQcBAQRo
# MGYwPQYIKwYBBQUHMAKGMWh0dHA6Ly9jcnQudXNlcnRydXN0LmNvbS9VVE5BZGRU
# cnVzdE9iamVjdF9DQS5jcnQwJQYIKwYBBQUHMAGGGWh0dHA6Ly9vY3NwLnVzZXJ0
# cnVzdC5jb20wDQYJKoZIhvcNAQEFBQADggEBABHJPeEF6DtlrMl0MQO32oM4xpK6
# /c3422ObfR6QpJjI2VhoNLXwCyFTnllG/WOF3/5HqnDkP14IlShfFPH9Iq5w5Lfx
# sLZWn7FnuGiDXqhg25g59txJXhOnkGdL427n6/BDx9Avff+WWqcD1ptUoCPTpcKg
# jvlP0bIGIf4hXSeMoK/ZsFLu/Mjtt5zxySY41qUy7UiXlF494D01tLDJWK/HWP9i
# dBaSZEHayqjriwO9wU6uH5EyuOEkO3vtFGgJhpYoyTvJbCjCJWn1SmGt4Cf4U6d1
# FbBRMbDxQf8+WiYeYH7i42o5msTq7j/mshM/VQMETQuQctTr+7yHkFGyOBkwggT+
# MIID5qADAgECAhArc9t0YxFMWlsySvIwV3JJMA0GCSqGSIb3DQEBBQUAMHoxCzAJ
# BgNVBAYTAkdCMRswGQYDVQQIExJHcmVhdGVyIE1hbmNoZXN0ZXIxEDAOBgNVBAcT
# B1NhbGZvcmQxGjAYBgNVBAoTEUNPTU9ETyBDQSBMaW1pdGVkMSAwHgYDVQQDExdD
# T01PRE8gVGltZSBTdGFtcGluZyBDQTAeFw0xOTA1MDIwMDAwMDBaFw0yMDA1MzAx
# MDQ4MzhaMIGDMQswCQYDVQQGEwJHQjEbMBkGA1UECAwSR3JlYXRlciBNYW5jaGVz
# dGVyMRAwDgYDVQQHDAdTYWxmb3JkMRgwFgYDVQQKDA9TZWN0aWdvIExpbWl0ZWQx
# KzApBgNVBAMMIlNlY3RpZ28gU0hBLTEgVGltZSBTdGFtcGluZyBTaWduZXIwggEi
# MA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQC/UjaCOtx0Nw141X8WUBlm7boa
# mdFjOJoMZrJA26eAUL9pLjYvCmc/QKFKimM1m9AZzHSqFxmRK7VVIBn7wBo6bco5
# m4LyupWhGtg0x7iJe3CIcFFmaex3/saUcnrPJYHtNIKa3wgVNzG0ba4cvxjVDc/+
# teHE+7FHcen67mOR7PHszlkEEXyuC2BT6irzvi8CD9BMXTETLx5pD4WbRZbCjRKL
# Z64fr2mrBpaBAN+RfJUc5p4ZZN92yGBEL0njj39gakU5E0Qhpbr7kfpBQO1NArRL
# f9/i4D24qvMa2EGDj38z7UEG4n2eP1OEjSja3XbGvfeOHjjNwMtgJAPeekyrAgMB
# AAGjggF0MIIBcDAfBgNVHSMEGDAWgBRkIoa2SonJBA/QBFiSK7NuPR4nbDAdBgNV
# HQ4EFgQUru7ZYLpe9SwBEv2OjbJVcjVGb/EwDgYDVR0PAQH/BAQDAgbAMAwGA1Ud
# EwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwQAYDVR0gBDkwNzA1Bgwr
# BgEEAbIxAQIBAwgwJTAjBggrBgEFBQcCARYXaHR0cHM6Ly9zZWN0aWdvLmNvbS9D
# UFMwQgYDVR0fBDswOTA3oDWgM4YxaHR0cDovL2NybC5zZWN0aWdvLmNvbS9DT01P
# RE9UaW1lU3RhbXBpbmdDQV8yLmNybDByBggrBgEFBQcBAQRmMGQwPQYIKwYBBQUH
# MAKGMWh0dHA6Ly9jcnQuc2VjdGlnby5jb20vQ09NT0RPVGltZVN0YW1waW5nQ0Ff
# Mi5jcnQwIwYIKwYBBQUHMAGGF2h0dHA6Ly9vY3NwLnNlY3RpZ28uY29tMA0GCSqG
# SIb3DQEBBQUAA4IBAQB6f6lK0rCkHB0NnS1cxq5a3Y9FHfCeXJD2Xqxw/tPZzeQZ
# pApDdWBqg6TDmYQgMbrW/kzPE/gQ91QJfurc0i551wdMVLe1yZ2y8PIeJBTQnMfI
# Z6oLYre08Qbk5+QhSxkymTS5GWF3CjOQZ2zAiEqS9aFDAfOuom/Jlb2WOPeD9618
# KB/zON+OIchxaFMty66q4jAXgyIpGLXhjInrbvh+OLuQT7lfBzQSa5fV5juRvgAX
# IW7ibfxSee+BJbrPE9D73SvNgbZXiU7w3fMLSjTKhf8IuZZf6xET4OHFA61XHOFd
# kga+G8g8P6Ugn2nQacHFwsk+58Vy9+obluKUr4YuMIIFYzCCBEugAwIBAgIRAI2m
# ZNIu7NSYIsW9qyNHzscwDQYJKoZIhvcNAQELBQAwfTELMAkGA1UEBhMCR0IxGzAZ
# BgNVBAgTEkdyZWF0ZXIgTWFuY2hlc3RlcjEQMA4GA1UEBxMHU2FsZm9yZDEaMBgG
# A1UEChMRQ09NT0RPIENBIExpbWl0ZWQxIzAhBgNVBAMTGkNPTU9ETyBSU0EgQ29k
# ZSBTaWduaW5nIENBMB4XDTE4MDcwODAwMDAwMFoXDTIwMDcwNzIzNTk1OVowga0x
# CzAJBgNVBAYTAkRFMQ4wDAYDVQQRDAUzMDUzOTEWMBQGA1UECAwNTmllZGVyc2Fj
# aHNlbjERMA8GA1UEBwwISGFubm92ZXIxEzARBgNVBAkMCkRyaWJ1c2NoIDIxJjAk
# BgNVBAoMHU5pY2sgSW5mb3JtYXRpb25zdGVjaG5payBHbWJIMSYwJAYDVQQDDB1O
# aWNrIEluZm9ybWF0aW9uc3RlY2huaWsgR21iSDCCASIwDQYJKoZIhvcNAQEBBQAD
# ggEPADCCAQoCggEBAKtzqVmbzG6U9EgZp6AVbKO4OSFibmp3NyPRlwuH4Yt58atm
# xNvxnCoPdfCE2wp5iZGV2zlyGDxjthWcdu1dVOG4vnU7rhboHmfYLNnJ4lUWtzoD
# IAYSdQ+XyijLLBxLoKH5rbzv0mai0tdYrBGcNu8bP84AZgTwxhan+4P/GRK+pblH
# AmgMrf97efppURJcPKP4x8yJlDIooYvtbm11wyeAClrU16veJ+TT4Z68PvNHIiYE
# fmaCXNerbYMGVswsF+GLshaIMV95ZuajE/ex36KYxGV6NH/+QD5DEd3GeW+m+DPq
# 6/I/NkO5/2bLdhhH8gclv+vCLlJNcsHgEfS5OrECAwEAAaOCAaswggGnMB8GA1Ud
# IwQYMBaAFCmRYP+KTfrr+aZquM/55ku9Sc4SMB0GA1UdDgQWBBQmgEarh9yZ7nLq
# 6/7q6RK2iwC7gzAOBgNVHQ8BAf8EBAMCB4AwDAYDVR0TAQH/BAIwADATBgNVHSUE
# DDAKBggrBgEFBQcDAzARBglghkgBhvhCAQEEBAMCBBAwRgYDVR0gBD8wPTA7Bgwr
# BgEEAbIxAQIBAwIwKzApBggrBgEFBQcCARYdaHR0cHM6Ly9zZWN1cmUuY29tb2Rv
# Lm5ldC9DUFMwQwYDVR0fBDwwOjA4oDagNIYyaHR0cDovL2NybC5jb21vZG9jYS5j
# b20vQ09NT0RPUlNBQ29kZVNpZ25pbmdDQS5jcmwwdAYIKwYBBQUHAQEEaDBmMD4G
# CCsGAQUFBzAChjJodHRwOi8vY3J0LmNvbW9kb2NhLmNvbS9DT01PRE9SU0FDb2Rl
# U2lnbmluZ0NBLmNydDAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuY29tb2RvY2Eu
# Y29tMBwGA1UdEQQVMBOBEWEubmlja0BuaWNrLWl0LmRlMA0GCSqGSIb3DQEBCwUA
# A4IBAQAH1ijdoz7SE9/WXcKSffHF2VW1ln/KGFpcMHZ4xKQvdG5IayaS9YFrrj5L
# jgNrDFx06qxlwySbIwULs0T00cqvfWZJUB9qQ2g5xVXXykbGZR6uaWwohiTKseSL
# wAw5VhjFYjWBZX20tZhmazCpZBpgAKZy77xGb3ri+NwNVw0L4Q0YXFPZY6JZKiVS
# 9uCfyfP+Y7iBUmQjfhm7FPmCPqYQ8tS3alZhpKzAMcdD3K2uxCtu6ap673KCb3FZ
# bmnO/sU3FIy2P9ekK0pSmGZotYon30cjPfKnJKT6Fq3rHlJ+t7j1yVomdv+ZTo/J
# VanfJHengqN9OwkXTK9CuaYpXxdhMIIF4DCCA8igAwIBAgIQLnyHzA6TSlL+lP0c
# t800rzANBgkqhkiG9w0BAQwFADCBhTELMAkGA1UEBhMCR0IxGzAZBgNVBAgTEkdy
# ZWF0ZXIgTWFuY2hlc3RlcjEQMA4GA1UEBxMHU2FsZm9yZDEaMBgGA1UEChMRQ09N
# T0RPIENBIExpbWl0ZWQxKzApBgNVBAMTIkNPTU9ETyBSU0EgQ2VydGlmaWNhdGlv
# biBBdXRob3JpdHkwHhcNMTMwNTA5MDAwMDAwWhcNMjgwNTA4MjM1OTU5WjB9MQsw
# CQYDVQQGEwJHQjEbMBkGA1UECBMSR3JlYXRlciBNYW5jaGVzdGVyMRAwDgYDVQQH
# EwdTYWxmb3JkMRowGAYDVQQKExFDT01PRE8gQ0EgTGltaXRlZDEjMCEGA1UEAxMa
# Q09NT0RPIFJTQSBDb2RlIFNpZ25pbmcgQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IB
# DwAwggEKAoIBAQCmmJBjd5E0f4rR3elnMRHrzB79MR2zuWJXP5O8W+OfHiQyESdr
# vFGRp8+eniWzX4GoGA8dHiAwDvthe4YJs+P9omidHCydv3Lj5HWg5TUjjsmK7hoM
# ZMfYQqF7tVIDSzqwjiNLS2PgIpQ3e9V5kAoUGFEs5v7BEvAcP2FhCoyi3PbDMKrN
# KBh1SMF5WgjNu4xVjPfUdpA6M0ZQc5hc9IVKaw+A3V7Wvf2pL8Al9fl4141fEMJE
# VTyQPDFGy3CuB6kK46/BAW+QGiPiXzjbxghdR7ODQfAuADcUuRKqeZJSzYcPe9hi
# KaR+ML0btYxytEjy4+gh+V5MYnmLAgaff9ULAgMBAAGjggFRMIIBTTAfBgNVHSME
# GDAWgBS7r34CPfqm8TyEjq3uOJjs2TIy1DAdBgNVHQ4EFgQUKZFg/4pN+uv5pmq4
# z/nmS71JzhIwDgYDVR0PAQH/BAQDAgGGMBIGA1UdEwEB/wQIMAYBAf8CAQAwEwYD
# VR0lBAwwCgYIKwYBBQUHAwMwEQYDVR0gBAowCDAGBgRVHSAAMEwGA1UdHwRFMEMw
# QaA/oD2GO2h0dHA6Ly9jcmwuY29tb2RvY2EuY29tL0NPTU9ET1JTQUNlcnRpZmlj
# YXRpb25BdXRob3JpdHkuY3JsMHEGCCsGAQUFBwEBBGUwYzA7BggrBgEFBQcwAoYv
# aHR0cDovL2NydC5jb21vZG9jYS5jb20vQ09NT0RPUlNBQWRkVHJ1c3RDQS5jcnQw
# JAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmNvbW9kb2NhLmNvbTANBgkqhkiG9w0B
# AQwFAAOCAgEAAj8COcPu+Mo7id4MbU2x8U6ST6/COCwEzMVjEasJY6+rotcCP8xv
# GcM91hoIlP8l2KmIpysQGuCbsQciGlEcOtTh6Qm/5iR0rx57FjFuI+9UUS1SAuJ1
# CAVM8bdR4VEAxof2bO4QRHZXavHfWGshqknUfDdOvf+2dVRAGDZXZxHNTwLk/vPa
# /HUX2+y392UJI0kfQ1eD6n4gd2HITfK7ZU2o94VFB696aSdlkClAi997OlE5jKgf
# cHmtbUIgos8MbAOMTM1zB5TnWo46BLqioXwfy2M6FafUFRunUkcyqfS/ZEfRqh9T
# TjIwc8Jvt3iCnVz/RrtrIh2IC/gbqjSm/Iz13X9ljIwxVzHQNuxHoc/Li6jvHBhY
# xQZ3ykubUa9MCEp6j+KjUuKOjswm5LLY5TjCqO3GgZw1a6lYYUoKl7RLQrZVnb6Z
# 53BtWfhtKgx/GWBfDJqIbDCsUgmQFhv/K53b0CDKieoofjKOGd97SDMe12X4rsn4
# gxSTdn1k0I7OvjV9/3IxTZ+evR5sL6iPDAZQ+4wns3bJ9ObXwzTijIchhmH+v1V0
# 4SF3AwpobLvkyanmz1kl63zsRQ55ZmjoIs2475iFTZYRPAmK0H+8KCgT+2rKVI2S
# XM3CZZgGns5IW9S1N5NGQXwH3c/6Q++6Z2H/fUnguzB9XIDj5hY5S6cxggRxMIIE
# bQIBATCBkjB9MQswCQYDVQQGEwJHQjEbMBkGA1UECBMSR3JlYXRlciBNYW5jaGVz
# dGVyMRAwDgYDVQQHEwdTYWxmb3JkMRowGAYDVQQKExFDT01PRE8gQ0EgTGltaXRl
# ZDEjMCEGA1UEAxMaQ09NT0RPIFJTQSBDb2RlIFNpZ25pbmcgQ0ECEQCNpmTSLuzU
# mCLFvasjR87HMA0GCWCGSAFlAwQCAQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKA
# AKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIODZMF65uM6jUbcMrkbLSTam
# WvsqdgP49qYInlPONqadMA0GCSqGSIb3DQEBAQUABIIBAANmB1iVaXdAv0AUmlSv
# 3ESJPw3ogmWMczWVTzeLlywhu2ouNGVTN1TdTzlXe6RThdaCZWS0l4Q/H3W/eoUR
# ZmxJo+hJeNL7UdomAvjW0hHfzDbBFflrtdGwQYeOBlgzr14jJBhjn+xMoFG0VCCr
# rQ2JfjMDPHmgy56IeXV7wD0WSyxX1GOK0g8v53IJYxLgtVGKm4K/dC7PqK9l1bsc
# EWSdnGCJLEwO9QgxN9t2ID1zwM57RbGszPt7mM18qqab8oXhyww9y9ZqEO6tvRIK
# B7CUumP9fxLBoUy9VHG2d9FpGs4y85jpkcTnuc8WEiT8F6yKa57KtcB1kcdYzlx8
# 9U2hggIoMIICJAYJKoZIhvcNAQkGMYICFTCCAhECAQEwgY4wejELMAkGA1UEBhMC
# R0IxGzAZBgNVBAgTEkdyZWF0ZXIgTWFuY2hlc3RlcjEQMA4GA1UEBxMHU2FsZm9y
# ZDEaMBgGA1UEChMRQ09NT0RPIENBIExpbWl0ZWQxIDAeBgNVBAMTF0NPTU9ETyBU
# aW1lIFN0YW1waW5nIENBAhArc9t0YxFMWlsySvIwV3JJMAkGBSsOAwIaBQCgXTAY
# BgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yMDAyMDcx
# NjQ1MzZaMCMGCSqGSIb3DQEJBDEWBBQ8dALp61qUMWeO+rl4cZ3BJXpEezANBgkq
# hkiG9w0BAQEFAASCAQBCQoXHTSE8PTh4KCHF9B4OJNl71s8I8GzSb+bWwtw8BD6r
# NOLXuN6sKVg9SVC2vMnAv5covGNIfqK8o0FwN5ab4ABwqClFO950naoNs0lYuFAX
# P1p6yCvRxRHnLRDNwi3w2qhFQ2ryoRhRMxDfNOA+1FJDUqPEkD41f2C4N9IoaR7S
# 96swirLj0b/OJFCcPko9Dus2diP99rSUKcmh3JXVv+3QnAeq2dgZfngiDzxAxK6C
# +9M98BWHgklaQQ8XddoGojRKJkd1zcIQ6FiHonIEtiayFqg6qWgGSSOx5vuxymxq
# p2qCgzdPWOUK3nC8+GE7sSHBvjjsg+4H2bLM6hyI
# SIG # End signature block
