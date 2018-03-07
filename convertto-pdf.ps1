<#
 # Helper functions to save common files to pdf format 
 # Author: PS Chakravarthy
#>

#https://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat.aspx

$pptType = "microsoft.office.interop.powerpoint.ppSaveAsFileType" -as [type]
$wdType = "microsoft.office.interop.word.WdSaveFormat" -as [type]
$xlType = "microsoft.office.interop.word.xlfileformat" -as [type]
$visioType = "microsoft.office.interop.visio.VisFixedFormatTypes" -as [type]

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

function Get-FileNames {
    param(
        $initialDirectory = $PSScriptRoot,
        $title = "Choose a file",
        $selectMultipleFiles = $false
    )
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.initialDirectory = $initialDirectory
    $fileDialog.title = $title
    $fileDialog.MultiSelect = $selectMultipleFiles
    $fileDialog.filter = "Office Files | *.ppt;*.pptx;*.doc;*.docx;*.xls;*.xlsx;*.vsd;*.vsdx;*.htm;*.html "
    $fileDialog.ShowDialog() | Out-Null
    $fileDialog.filenames
}

function Show-Message {
    param($msg, $type = "I")

    $notifyIcon = New-Object System.Windows.Forms.NotifyIcon 
    $notifyIcon.Icon = "$PSScriptRoot\ppt.ico"
    $notifyIcon.BalloonTipIcon = "Info" 
    $notifyIcon.BalloonTipText = $msg
    $notifyIcon.BalloonTipTitle = "PPT-UTILS Conversion" 
    $notifyIcon.Visible = $True 
    $notifyIcon.ShowBalloonTip(2000)
    $notifyIcon.Dispose()
}

<#
 # Apply given PPT template to supplied source file. 
 # Assumes source presentations are using standard layouts in a presentation.
 # Quality of new presentation is dependent on proper use of slide layouts.
#>

function Apply-PPT-Template {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,Position=0)] $src, 
        [string] $dest = $src.replace(".pptx", "-new.pptx"),
        [string] $template = $PPT_TEMPLATE
     )
 
    Write-Verbose "Applying $template to $src" 
    $presentationApp = New-Object -ComObject powerpoint.application
    $presentationApp.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $presentation = $presentationApp.Presentations.Open($src)
    $presentation.ApplyTemplate($template)

    $presentation.slides | 
    ForEach-Object -Begin {$count = 1 }  `
        -Process { 
            $currentSlide = $_;
            Clean-Title-Slides -slide $currentSlide;
            $count++ } `
        -End { "" }


    Write-Verbose "Save output file $dest"
    $presentation.SaveAs($dest, [microsoft.office.interop.powerpoint.ppSaveAsFileType]::ppSaveAsDefault)
    $presentation.Close()
    $presentationApp.Quit()

    if (Test-Path $dest) {
        Write-Output "Created $dest"
    }
}


function Convert-Word-to-PDF {
    param($srcFile)

    $dest = $srcFile -replace "(.docx)|(.doc)", ".docx.pdf"
   $dest =  $dest.insert($dest.LastIndexOf("\"), "\out")
    Write-Output "Creating PDF for $srcFile"
    $app = New-Object -ComObject word.application
    $app.Visible = $fal
    $doc = $app.Documents.Open($srcFile)
    $doc.SaveAs([ref]$dest, [ref] $wdType::wdFormatPDF)
    $doc.Close()
    $app.Quit()

    if (Test-Path $dest) {
        Write-Output "Created $dest"
    }
}

function Convert-Excel-to-PDF {
    param($srcFile)

    $dest = $srcFile -replace "(.xlsx)|(.xls)", ".xlsx.pdf"
   $dest =  $dest.insert($dest.LastIndexOf("\"), "\out")

    Write-Output "Creating PDF for $srcFile"
    $app = New-Object -ComObject excel.application
    $app.Visible = $false
    $doc = $app.workbooks.Open($srcFile)
    $doc.saved = $true
    $doc.ExportAsFixedFormat($xlType::xlTypePDF, $dest)
    $doc.close()
    $app.Quit()

    if (Test-Path $dest) {
        Write-Output "Created $dest"
    }
}

function Convert-PPT-to-PDF {
    param($srcFile)

    $dest = $srcFile -replace "(.pptx)|(.ppt)", ".pptx.pdf"
   $dest =  $dest.insert($dest.LastIndexOf("\"), "\out")

    Write-Output "Creating PDF for $srcFile"
    $app = New-Object -ComObject powerpoint.application
    $doc = $app.Presentations.Open($srcFile)
    $doc.SaveAs($dest, $pptType::ppSaveAsPDF)
    $doc.close()
    $app.Quit()

    if (Test-Path $dest) {
        Write-Output "Created $dest"
    }
}

function Convert-HTML-to-PDF {
    param($srcFile)

    $dest = $srcFile -replace "(.html)|(.htm)", ".html.pdf"
   $dest =  $dest.insert($dest.LastIndexOf("\"), "\out")


    set-alias chrome  "c:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    $options = "--headless --disable-gpu --enable-local-file-accesses --print-to-pdf='$($dest)' '$($srcFile)'"
    $p = "chrome $options"

    Invoke-Expression $p
    Stop-Process -Name Chrome

}

function Convert-Visio-to-PDF {
    param($srcFile)

    $dest = $srcFile -replace "(.vsdx)|(.vsd)", ".vsdx.pdf"
   $dest =  $dest.insert($dest.LastIndexOf("\"), "\out")
    Write-Output "Creating PDF for $srcFile"
    $app = New-Object -ComObject visio.application
    $doc = $app.documents.Open($srcFile)
    $doc.ExportAsFixedFormat([microsoft.office.interop.visio.VisFixedFormatTypes]::visFixedFormatPDF, $dest, [microsoft.office.interop.visio.VisDocExIntent]::visDocExIntentPrint,         [microsoft.office.interop.visio.VisPrintOutRange]::visPrintAll)
    $doc.close()
    $app.Quit()

    if (Test-Path $dest) {
        Write-Output "Created $dest"
    }
}


function ConvertTo-PDF {
    Get-FileNames -selectMultipleFiles $true | ForEach-Object -Process {
        $file = $_
        $ext = $file.substring($_.lastindexof(".") + 1)

        switch -Wildcard ($ext) {
            "doc*" { Convert-Word-to-PDF -srcFile $file; break}
            "xls*" { Convert-Excel-to-PDF -srcFile $file; break}
            "ppt*" { Convert-PPT-to-PDF -srcFile $file; break}
            "vsd*" { Convert-Visio-to-PDF -srcFile $file; break}
            "htm*" { Convert-HTML-to-PDF -srcFile $file; break}
        }
    }
    
}

ConvertTo-PDF