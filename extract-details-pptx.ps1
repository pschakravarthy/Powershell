<#
.SYNPOSIS
    Helper functions to extract links and notes from a PowerPoint
    Presentation. Creates a org-mode file which can be processed
    to produce other formas.

.NOTES
    Name: extract-links-pptx
    Author: PS Chakravarthy
#>

$ppType = "microsoft.office.interop.powerpoint.ppPlaceholderType" -as [type]

function Emit-Notes {
    param($slides)

    Write-Output ""
    Write-Output "## Slide Notes"
    Write-Output "|Slide       |   Notes"
    Write-Output "|------------| --------------------------------------" 

    foreach ($slide in $slides) { 
        if ($slide.HasNotesPage) {
            foreach ($shape in $slide.notespage.shapes) {
                if (($shape.placeholderformat.type -eq $ppType::ppPlaceholderBody) -and $shape.hasTextFrame -and $shape.TextFrame.hasText) {
                    Write-Output "| $($slide.SlideIndex)     | $($shape.TextFrame.TextRange.Text)"
                }
            }
        }
    }
}

function Emit-Links {
    param($slides)

    Write-Output ""
    Write-Output "## Hyperlinks in Presentation"
    Write-Output "|Slide       |   URL"
    Write-Output "|------------| --------------------------------------" 

    foreach ($slide in $slides) { 
	    foreach ($link in $slide.hyperlinks) { 
            $href = $link.Address
            $href = $href.replace(" ", "%20")
	        Write-Output "| $($slide.SlideIndex)| [[$href]]" 
	    }
    }
}

function Get-Presentation-Details {
    param ($presentation)
    
    $slides = $presentation.Slides

    Write-Output '% Title: Presentation Details
    Write-Output ""
    Write-Output "## Presentation Details"
    Write-Output ""
    Write-Output "|Name | [[$($presentation.FullName)][ $($presentation.Name)]]"
    Write-Output "|Template |  $($presentation.templatename)"
    Write-Output "|Slides | $($presentation.Slides.count)"
    Write-Output " "

    Emit-Links -slides $slides
    Emit-Notes -slides $slides
}

function Extract-Details {
    param($sourcePPT,
    $outputFile)

    if (Test-Path -path $outputFile) {
	    Remove-Item $outputFile -Force
    }

    $app = New-Object -ComObject powerpoint.application  
    $presentation = $app.Presentations.open($sourcePPT)
    
    Get-Presentation-Details -presentation $presentation | Out-File -Filepath $outputFile -Force
    
    $app.ActivePresentation.Close()
    
    $app.Quit()
    $app = $null
}

