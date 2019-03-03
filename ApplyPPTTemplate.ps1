<#
 # Template Utils to manipulate PowerPoint 
 # Author: PS Chakravarthy
#>

#https://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat.aspx

$pptType = "microsoft.office.interop.powerpoint.ppSaveAsFileType" -as [type]

<#
 # Apply given PPT template to supplied source file. 
 # Assumes source presentations are using standard layouts in a presentation.
 # Quality of new presentation is dependent on proper use of slide layouts.
#>

function Apply-PPT-Template {
    param(
        [string] $src, 
        [string] $dest = $src.replace(".pptx", "-new.pptx"),
        [string] $template = $PPT_TEMPLATE
     )
    Write-Output "Applying $template to $src"
    $presentationApp = New-Object -ComObject powerpoint.application
    $presentationApp.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $presentation = $presentationApp.Presentations.Open($src)
    $presentation.ApplyTemplate($template)
    $presentation.SaveAs($dest, $pptType::ppSaveAsDefault)
    $presentation.Close()
    $presentationApp.Quit()

    if (Test-Path $dest) {
        Write-Output "Created $dest"
    }
}
