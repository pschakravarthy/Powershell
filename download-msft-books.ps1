############################################################### 
# Eric Ligmans Amazing Free Microsoft eBook Giveaway 
# https://blogs.msdn.microsoft.com/mssmallbiz/2017/07/11/largest-free-microsoft-ebook-giveaway-im-giving-away-millions-of-free-microsoft-ebooks-again-including-windows-10-office-365-office-2016-power-bi-azure-windows-8-1-office-2013-sharepo/
# Link to download list of eBooks 
# http://ligman.me/2sZVmcG 
# Thanks David Crosby for the template (https://social.technet.microsoft.com/profile/david%20crosby/)
############################################################### 
$dest = "C:\users\pavanis\documents\Downloads\ebooks\" 
 
# Download the source list of books 
$downLoadList = "http://ligman.me/2sZVmcG" 
$bookList = Invoke-WebRequest $downLoadList 
 
# Convert the list to an array 
[string[]]$books = "" 
$books = $bookList.Content.Split("`n") 
# Remove the first line - it's not a book 
$books = $books[1..($books.Length -1)] 
$books # Here's the list 
 
# Download the books 
foreach ($book in $books) { 
    $hdr = Invoke-WebRequest $book -Method Head 
    $title = $hdr.BaseResponse.ResponseUri.Segments[-1] 
    $title = [uri]::UnescapeDataString($title) 
    $saveTo = $dest + $title 

    if (Test-Path -path $saveTo ) {
        Write-Output "File exists: $saveTo"
    }
    else {
        Invoke-WebRequest $book -OutFile $saveTo
    } 
} 
