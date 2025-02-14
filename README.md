# 📄 PowerShell Script for Counting PDF Pages in SharePoint Online

## 🎯 Overview
This PowerShell script automatically **counts the number of pages in PDF files** stored in a SharePoint Online document library and updates a custom column named `PageCount`. It uses **PnP PowerShell** and **PdfSharpCore** to process and analyze PDFs efficiently.

### 🚀 **Key Features**
✅ Automatically **counts pages** in PDF files stored in SharePoint.  
✅ Uses **PnP PowerShell** for seamless SharePoint Online connectivity.  
✅ Updates SharePoint metadata with the **PageCount** value.  
✅ Handles errors gracefully and cleans up temporary files.  

---
## 🛠 Prerequisites
Before running the script, ensure the following requirements are met:

### 🔧 **Required Tools**
1. **PnP PowerShell Module** (If not installed, run):
   ```powershell
   Install-Module PnP.PowerShell -Scope CurrentUser
   ```
2. **PdfSharpCore Library** for PDF processing.
   - Download and extract `PdfSharpCore.dll` and place it in the correct path.
3. **SharePoint Online Access** with edit permissions.
4. **PageCount Column** in SharePoint.
   - Create a **Number column** in the document library and name it `PageCount`.

---
## 📝 PowerShell Script

```powershell
# Connect to SharePoint Online
Connect-PnPOnline -Url "https://futurrizoninterns.sharepoint.com/sites/mentalhealthcarewebapplication1" -UseWebLogin

# Load PdfSharpCore Library
Add-Type -Path "E:\Work FT\pdfsharpcore.1.3.65\lib\net8.0\PdfSharpCore.dll"

# Function to Get PDF Page Count
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
        Write-Host "❌ Error processing file: $pdfPath - $_"
        return $null
    }
}

$libraryName = "CustomDocumentLibrary"

# Get all PDF files in the SharePoint document library
$files = Get-PnPListItem -List $libraryName -Fields "FileRef", "ID" | Where-Object { $_["FileRef"] -like "*.pdf" }

foreach ($file in $files) {
    $fileUrl = $file["FileRef"]
    $localPath = "$env:TEMP\" + [System.IO.Path]::GetFileName($fileUrl)

    # Download the PDF file
    Get-PnPFile -Url $fileUrl -Path $env:TEMP -FileName ([System.IO.Path]::GetFileName($fileUrl)) -AsFile -Force

    # Get page count
    $PageCount = Get-PDFPageCount -pdfPath $localPath

    if ($pageCount -ne $null) {
        # Update SharePoint metadata
        Set-PnPListItem -List $libraryName -Identity $file.Id -Values @{"PageCount" = $pageCount}
        Write-Host "✅ Updated $fileUrl with $pageCount pages."
    }

    # Cleanup temporary file
    Remove-Item -Path $localPath -Force
}
```

---
## 🚀 How to Use the Script

1. **Modify the script**:
   - Update `$SiteURL` with your SharePoint Online site.
   - Change `$libraryName` to your document library name.
2. **Ensure you have `PageCount` column** in your SharePoint library.
3. **Run the script** in PowerShell:
   ```powershell
   .\YourScriptName.ps1
   ```
4. **Check SharePoint Online**:
   - The `PageCount` column will be updated with the number of pages for each PDF file.

---
## 🔥 Troubleshooting & FAQs

### ❓ **1. What if `PageCount` column is missing?**
📌 Create a **Number column** in your SharePoint library and name it `PageCount`.

### ❓ **2. What if the script fails to connect?**
📌 Ensure PnP PowerShell is installed and use `-UseWebLogin` for authentication.

### ❓ **3. Can I use this for other file types?**
📌 No, this script is specifically designed for **PDF files**.

### ❓ **4. What if `Error processing file` appears?**
📌 Ensure the **PdfSharpCore.dll** path is correct and the PDFs are not corrupted.

---
## 🔍 What is this script for?
This script is making it easy to find out how many pages are in your PDF files in SharePoint Online. Its a great way to keep track of your documents and make it easier to manage your library.:
✅ "How to count PDF pages in SharePoint using PowerShell"
✅ "PnP PowerShell extract PDF page count"
✅ "Update SharePoint Online metadata automatically"
✅ "PowerShell script to count pages in PDF files"

If you're looking for an easy way to **count PDF pages in SharePoint Online**, this script is the perfect solution! 🚀📂

---
