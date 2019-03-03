<#
.SYNOPSIS
Helper functions to support other scripts. 

.DESCRIPTION
Helper functions to allow scripts to zip/unzip, make-pdf using Chrome, and write to application log.

.NOTES
    Author: PS Chakravarthy
#>
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Unzip
{
    param([string]$zipfile, [string]$outpath)

    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
}

function Zip
{
    param([string]$zipfile, [string]$srcpath)

    [System.IO.Compression.ZipFile]::CreateFromDirectory($srcpath, $zipfile)
}

# Helper function to write to application log
function Write-Log {
    param([string]$scriptName,
	  [string]$entryType="Information",
	  [string]$msg, 
	  [string]$eventId=1000)

    # Create log source using the script name
    $logType = "Application"
    $logExists = ([System.Diagnostics.EventLog]::Exists($logType) -and [System.Diagnostics.EventLog]::SourceExists($scriptName) )
    if (! $logExists) {
	New-EventLog -LogName $logType -Source $scriptName | out-null
    }

    # Write the log event
    Write-EventLog -LogName $logType -Source $scriptName -EntryType $entryType  -EventID $eventId -Message $msg
}

function make-pdf {
    param($html)
    $pdf = $html -replace ".html",".pdf"
    $pdfParams =  '--headless --disable-gpu --enable-local-file-accesses --print-to-pdf="$($pdf)" "$($html)"'
    $chrome = 'c:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
    & $chrome --headless --disable-gpu --enable-local-file-accesses --print-to-pdf="$($pdf)" "$($html)"
}

<#
.DESCRIPTION
Sometimes need to cleanup html files for conversion using Pandoc. This function is meant to help cleanup such files.
#>
function Cleanup-Utf {
    param(
        $path = "",
        $filter = "*.html"
    )

    $files= get-childitem -path $path -filter $filter
    foreach ($f in $files) {
        $content = get-content -path $f.FullName
        write-output "Processing $($f.FullName)"
        Set-Content -Path $f.FullName -Encoding UTF8 -Value $content
        }

}
