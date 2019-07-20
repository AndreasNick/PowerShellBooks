function Add-PageNumbers
        {
            param
            (
                [System.IO.FileInfo] $InputPdfDocument,
                [System.IO.FileInfo] $OutputPdfDocument,
                [Bool] $ScipFirstPage = $true
            )
           
            if (test-path "$OutputPdfDocument") { Remove-Item "$OutputPdfDocument"  }

            $reader = New-Object iTextSharp.text.pdf.PdfReader -ArgumentList $InputPdfDocument.fullname
            $fileStream = New-Object System.IO.FileStream($OutputPdfDocument, [System.IO.FileMode]::OpenOrCreate)
            $stamper = New-Object iTextSharp.text.pdf.PdfStamper -ArgumentList @($reader, $fileStream)
            
            $box = $reader.GetPageSize(1)
            
            $Start = 1
            if($ScipFirstPage){

                $Start=2
            }

            for ($i = $Start; $i -le $reader.NumberOfPages; $i++)
            {
                
                $canvas =  $stamper.GetOverContent($i)
                $page = $stamper.GetImportedPage($reader, $i)
                
                $font= [iTextSharp.text.pdf.BaseFont]::CreateFont([iTextSharp.text.pdf.BaseFont]::HELVETICA,[iTextSharp.text.pdf.BaseFont]::CP1252,[iTextSharp.text.pdf.BaseFont]::NOT_EMBEDDED)
                
                $canvas.SetColorFill([iTextSharp.text.BaseColor]::DARK_GRAY);
                $canvas.BeginText();
                $canvas.SetFontAndSize($font, 11);
                
                #Write-host $( $box.Height)
                $canvas.ShowTextAligned([iTextSharp.text.pdf.PdfContentByte]::ALIGN_CENTER, $(<#"page  " +#> $i), $box.Width/2 , 30, 0)
                $canvas.EndText();
                $canvas.AddTemplate($page, 0, 0);
                
            }
            $stamper.Close() 
           $reader.Close()
           
           $fileStream.Close()
        }