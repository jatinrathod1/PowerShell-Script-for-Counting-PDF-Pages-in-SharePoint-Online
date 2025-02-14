# Page Count 

Connect-PnPOnline -Url "https://futurrizoninterns.sharepoint.com/sites/mentalhealthcarewebapplication1" -UseWebLogin

Add-Type -Path "E:\Work FT\pdfsharpcore.1.3.65\lib\net8.0\PdfSharpCore.dll"


function Get-PDFPageCount {
    param ([string]$pdfPath)

    try {
        # Load PdfSharpCore and open the document
        $document = [PdfSharpCore.Pdf.IO.PdfReader]::Open($pdfPath, [PdfSharpCore.Pdf.IO.PdfDocumentOpenMode]::Import)

        # Get the number of pages
        $pageCount = $document.PageCount
        $document.Close()

        return $pageCount
    }
    catch {
        Write-Host "Error processing file: $pdfPath - $_"
        return $null
    }
}

$libraryName = "CustomDocumentLibrary"

# Get all PDF files in the document library
$files = Get-PnPListItem -List $libraryName -Fields "FileRef", "ID" | Where-Object { $_["FileRef"] -like "*.pdf" }

foreach ($file in $files) {
    $fileUrl = $file["FileRef"]
    $localPath = "$env:TEMP\" + [System.IO.Path]::GetFileName($fileUrl)

    # Download the PDF file
    Get-PnPFile -Url $fileUrl -Path $env:TEMP -FileName ([System.IO.Path]::GetFileName($fileUrl)) -AsFile -Force

    # Get page count
    $PageCount = Get-PDFPageCount -pdfPath $localPath

    if ($pageCount -ne $null) {
        # Update SharePoint column
        Set-PnPListItem -List $libraryName -Identity $file.Id -Values @{"pageCount" = $pageCount}
        Write-Host "Updated $fileUrl with $pageCount pages."
    }

    # Clean up
    Remove-Item -Path $localPath -Force
}