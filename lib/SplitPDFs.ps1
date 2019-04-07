#DEV ON JESTER Version 1.0.2

#Change Log
<#
    2014.4.26 - Added support for multiple leves of bookmarks
                Adjusted display when a page is "guessed" to only show when actually duplicated
                Added slight twist on control structre to not use page lables with numerics only. ie if page 1 -eq 1, to not use it
    2014.4.25 - Added first ability to use bookmarks not page labels
    2014.4.20 - Added "$pdfname-Extracted" to output file name to avoid naming conflicts
    2017.4.19 - Added exit if no page labels exist 
    
#>
#Description
<#
    Author: Ben Wheat
    Date:   2017.04.25

    This script uses the iTextSharp c# library to split a pdf into seperate pages and rename them
    based on the pdf lables (page names), if existing. This library powers the copy of each page. The library can also
    split mult pages if necessary. 

    The script will take into account Bookmarks (outlines) and identify any page numbers referenced.This feature will be 
    stronger the more pdfs i'm able to adjust too. The script will prefx 'Duplicate-' to any pages with no page reference
    in the bookmarks. 

    I used many References, search the tag REF to find all links. 
    REF:
    Source for library here: https://sourceforge.net/projects/itextsharp/
    Index for library: http://afterlogic.com/mailbee-net/docs-itextsharp/Index.html
#>
#REF: https://www.codeproject.com/Articles/559380/Splitting-and-Merging-PDF-Files-in-Csharp-Using-i

Write-Host "Launching Split PDFs Utility...."

#TODO: $use some kind of env variable to get the root of the folder
#these needs to be set manually right now
$dir = "\\Jester\Software\Miscellaneous\Scripts\SplitPDFs"

#REF: https://www.codeproject.com/Articles/559380/Splitting-and-Merging-PDF-Files-in-Csharp-Using-iT
#REF: for C#
Add-Type -ReferencedAssemblies "$dir\lib\itextsharp.dll" -typedefinition @"
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;

// CLASS DEPENDS ON iTextSharp: http://sourceforge.net/projects/itextsharp/

namespace iTextTools
{
    public class PdfExtractorUtility
    {
        public void ExtractPages(string sourcePdfPath, 
            string outputPdfPath, int[] extractThesePages)
        {
            PdfReader reader = null;
            Document sourceDocument = null;
            PdfCopy pdfCopyProvider = null;
            PdfImportedPage importedPage = null;

            try
            {
                // Intialize a new PdfReader instance with the 
                // contents of the source Pdf file:
                reader = new PdfReader(sourcePdfPath);

                // For simplicity, I am assuming all the pages share the same size
                // and rotation as the first page:
                sourceDocument = new Document(reader.GetPageSizeWithRotation(extractThesePages[0]));

                // Initialize an instance of the PdfCopyClass with the source 
                // document and an output file stream:
                pdfCopyProvider = new PdfCopy(sourceDocument,
                    new System.IO.FileStream(outputPdfPath, System.IO.FileMode.Create));

                sourceDocument.Open();

                // Walk the array and add the page copies to the output file:
                foreach (int pageNumber in extractThesePages)
                {
                    importedPage = pdfCopyProvider.GetImportedPage(reader, pageNumber);
                    pdfCopyProvider.AddPage(importedPage);
                }
                sourceDocument.Close();
                reader.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public void ExtractPages(string sourcePdfPath, string outputPdfPath, 
            int startPage, int endPage)
        {
            PdfReader reader = null;
            Document sourceDocument = null;
            PdfCopy pdfCopyProvider = null;
            PdfImportedPage importedPage = null;

            try
            {
                // Intialize a new PdfReader instance with the contents of the source Pdf file:
                reader = new PdfReader(sourcePdfPath);

                // For simplicity, I am assuming all the pages share the same size
                // and rotation as the first page:
                sourceDocument = new Document(reader.GetPageSizeWithRotation(startPage));

                // Initialize an instance of the PdfCopyClass with the source 
                // document and an output file stream:
                pdfCopyProvider = new PdfCopy(sourceDocument, 
                    new System.IO.FileStream(outputPdfPath, System.IO.FileMode.Create));

                sourceDocument.Open();

                // Walk the specified range and add the page copies to the output file:
                for (int i = startPage; i <= endPage; i++)
                {
                    importedPage = pdfCopyProvider.GetImportedPage(reader, i);
                    pdfCopyProvider.AddPage(importedPage);
                }
                sourceDocument.Close();
                reader.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public void ExtractPage(string sourcePdfPath, string outputPdfPath, 
            int pageNumber)
        {
            PdfReader reader = null;
            Document document = null;
            PdfCopy pdfCopyProvider = null;
            PdfImportedPage importedPage = null;

            try
            {
                // Intialize a new PdfReader instance with the contents of the source Pdf file:
                reader = new PdfReader(sourcePdfPath);

                // Capture the correct size and orientation for the page:
                document = new Document(reader.GetPageSizeWithRotation(pageNumber));

                // Initialize an instance of the PdfCopyClass with the source 
                // document and an output file stream:
                pdfCopyProvider = new PdfCopy(document, 
                    new System.IO.FileStream(outputPdfPath, System.IO.FileMode.Create));

                document.Open();

                // Extract the desired page number:
                importedPage = pdfCopyProvider.GetImportedPage(reader, pageNumber);
                pdfCopyProvider.AddPage(importedPage);
                document.Close();
                reader.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

            }
        }
    }
}
"@

#open dialog window to select PDF and get path to it
Function get-Filename($initialDir) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.forms") | Out-Null

    $OpenFileDiag = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDiag.InitialDirectory = $initialDir
    $OpenFileDiag.Filter = "PDF (*.pdf)| *.pdf" #limit to only PDFs
    $OpenFileDiag.ShowDialog() | Out-Null
    $OpenFileDiag.FileName
}
#save the path of the file
$sourcePDFPath = get-Filename -initialDir $env:USERPROFILE\Desktop

#import dll
Add-Type -Path $dir\lib\itextsharp.dll

#test document opening
try{
    $document = New-Object iTextSharp.text.pdf.PdfReader -ArgumentList $sourcePDFPath -ErrorAction Stop
    $docTotalPages = $document.NumberOfPages
    $arrayofLabels = [iTextSharp.Text.pdf.pdfpagelabels]::getpagelabels($document)
}catch{
    Write-Host "Error importing PDF file"
    pause
    exit
    #TODO: add error correction opportunity
}

#Drive the script based off existing page lables or not
#TODO: select method better
if(($arrayofLabels.Length -ne $docTotalPages) -or ($arrayofLabels[0] -eq "1")){
    #below will stop this if block if necessary.
    #TODO: Add switch option of partial page labels and do same idea as bookmarks
    #TODO: add ability to make do with partial of each (bookmarks and page labels)
    <#
    #go ahead an exit of no page labels are found
    if($arrayofLabels.length -eq 0){
        Write-Host "This PDF has no page labels! Exiting now"
        pause
        exit
    }
    #>
    #REF: https://www.codeproject.com/tips/480622/extract-bookmark-from-pdf-file
    ##########################################################################
    #Set path labels and create folder
    $folderName = [regex]::Split($sourcePDFPath,'.pdf$') #remove .pdf tag
    $pdfName = %{$split = [regex]::Match($folderName,'[\-\w\.\ ]*$'); $split[0].Value} #select the name up to the \ in the path
    $folderName = "Extracted-" + $pdfName
    #create save to folder in root where sourcePDF is located called "Extracted-$pdfname"
    [string]$destFolder = [regex]::Split($sourcePDFPath,'[\-\w\d\s]*.pdf$') -join $folderName
    #TODO: why white space at end??
    $destFolder =  [regex]::Replace($destFolder,'[\s]$','') #remove white space at end
    #TODO: this will error if folder already exists unless error action exists
    New-Item $destFolder -type Directory -ErrorAction SilentlyContinue
    ##########################################################################
    #memory stream for export XML
    $ms = New-Object System.IO.MemoryStream
    $arrayofBookmarks = [iTextSharp.text.pdf.SimpleBookmark]::GetBookmark($document)
    [iTextSharp.text.pdf.SimpleBookmark]::ExportToXML($arrayofBookmarks,$ms,"ISO8859-1", $true)
    $sr = New-Object System.IO.StreamReader($ms)
    $ms.Position = 0
    #exportXML Path to extracted folder to avoid confilcts with multi users
    $outXML = "$destFolder\bookmarks.xml"
    #export xml
    $sr.ReadToEnd() | Out-File -FilePath $outXML
    #import xml file
    [xml]$inBookmarksXML = Get-Content -Path $outXML
    #TODO try catch here to get errors if this doesn't work with other pdfs
    ##########################################################################
    #use the fact the result is a null value if no bookmarks are found to get to the correct bookmark depth and return that
    if($inBookmarksXML.Bookmark.Title.Title.Title.Title -eq $null){
        if($inBookmarksXML.Bookmark.Title.Title.Title -eq $null){
            if($inBookmarksXML.Bookmark.Title.Title-eq $null){
                if($inBookmarksXML.Bookmark.Title -eq $null){
                    Write-Host "Bookmarks not found. Exiting Utility.."
                    exit
                }else{
                    $listofBookmarks = $inBookmarksXML.Bookmark.Title
                }
            }else{
                $listofBookmarks = $inBookmarksXML.Bookmark.Title.Title
            }
        }else{
            $listofBookmarks = $inBookmarksXML.Bookmark.Title.Title.Title
        }
    }else{
        $listofBookmarks = $inBookmarksXML.Bookmark.Title.Title.Title.Title
    }
    #make array from the XML formatted bookmark file
    #$listofBookmarks = $inBookmarksXML.Bookmark.Title.Title.Title #this might not work on all pdfs..................
    ##########################################################################
    if($listofBookmarks.Count -gt $docTotalPages){
        #More than necessary
    }
    #Process to get names of pages and numbers in order
    #Initialize array slightly large, so each page is at the actualy page number
    #TODO adjust to standard array not oversized one
    $arraytotalpages = @($null) * ($docTotalPages+1)
    #Go through what pages are known 
    foreach($item in $listofBookmarks){
        #Get page number from array list of bookmarks using REGEX
        #Other items not number will be discarded
        $index = [int][regex]::Match($item.Page, '\d*').Value
        #make new object to house all the complete number of page names
        $pdfPageObj = [PsCustomObject]@{
            Page = [int][regex]::Match($item.Page, '\d*').Value
            Title = [string]$item.'#text'
         }
         #assign accurate page to the array of complete pages. 
         #i.e. page 3 from XML output is going to index #3 in this array
         $arraytotalpages[$index] += $pdfPageObj
    }
    $index = 0
    #set skipped pages
    $guessedPages = @()
    #Go through each looking for $null items (missed pages)
    #Add any missing with a duplicate of the index's previous
    $guessHappend = $false
    foreach($item in $arraytotalpages){
        #remember to skip the first as it will always be null technically
        if(($item -eq $null) -and ($index -ne 0)){
            #TODO: clean this write-host up when other output to user is cleaned up
           Write-Host "Adding page at $index"
           $guessHappend = $true
           $pdfPageObj = [PsCustomObject]@{
                Page = $index
                Title = "DUPLICATE-" + $arraytotalpages[$index-1].Title
            }
            #make a guessed pages array to display to user later
            $guessedPages += @($index)
            $arraytotalpages[$index] = $pdfPageObj
        }
        $index++
    }
    $page = 1
    #loop through all papes
    while($page -le $docTotalPages){
        Write-Progress -Activity "Spliting PDF... ['ctrl+c to' force stop]" -PercentComplete ($page / $docTotalPages * 100) -CurrentOperation "Page: $page of $docTotalPages."
        #set save to dest
        #Regex to remove disallowed characters from string of path
        $pdfName = [regex]::Replace($arraytotalpages[$page].Title, '[^\w\-\.\ ]', '')
        $savetoPath = $destFolder + "\" + $pdfName + ".pdf"
        $newPDF = [iTextTools.PdfExtractorUtility]::new()
         #count here so above path is accurate in getting label, and below is accurate on what page
        try{
            $newPDF.ExtractPage($sourcePDFPath, $savetoPath, $page)
        }catch{
            Write-Host "Error with page #$page" + $_.Exception.Message
            pause
            $document.Close()
            exit
        }
        $page++
    }
    #remove the XML file created early
    Remove-Item $outXML
    #let user know if any items were guessed upon
    #TODO get this cleaned up
    if($guessHappend -eq $true){
        Write-Host "There was not a complete list of bookmarks in this pdf. Guessed items will have the prefix 'Duplicate-'"
        Write-Host "The pages without accurate bookmarks are" -NoNewline  
        foreach($page in $guessedPages) {Write-host " " $page -NoNewline}
    }
    Write-Host ''
    Write-Host "Extracted files are located in $destFolder"
    #Close files and streams
    $document.Close()
    $sr.Close()
    $ms.Close()
    ########################
    ####end if, now else####
    ########################
    }else{ 
    #files have labels for each page
    ##########################################################################
    #Set path labels
    $folderName = [regex]::Split($sourcePDFPath,'.pdf$') #remove .pdf tag
    $pdfName = %{$split = [regex]::Match($folderName,'[\-\w\.\ ]*$'); $split[0].Value} #select the name up to the \ in the path
    #$pdfName =  [regex]::Split($pdfName,'[\s]$') #remove white space at end of pdf name
    $folderName = "Extracted-" + $pdfName
    #$folderName =  [regex]::Split($folderName,'[\s]$') #remove white space at end 
    #create save to folder in root where sourcePDF is located called "Extracted"
    [string]$destFolder = [regex]::Split($sourcePDFPath,'[\-\w\d\s]*.pdf$') -join $folderName
    $destFolder =  [regex]::Replace($destFolder,'[\s]$','') #remove white space at end
    New-Item $destFolder -type Directory
    ##########################################################################
    $page = 0
    #loop through all papes
    while($page -lt $docTotalPages){
        Write-Progress -Activity "Spliting PDF... ['ctrl+c to' force stop]" -PercentComplete ($page / $docTotalPages * 100) -CurrentOperation "Page: $page of $docTotalPages."
        #set save to dest
        #Regex to remove disallowed characters from string of path
        $pdfName = [regex]::Replace($arrayofLabels[$page], '[^\w\-\.\ ]', '')
        $savetoPath = $destFolder + "\" + $pdfName + ".pdf"
        $newPDF = [iTextTools.PdfExtractorUtility]::new()
        #count here so above path is accurate in getting label, and below is accurate on what page
        $page++ 
        try{
            $newPDF.ExtractPage($sourcePDFPath, $savetoPath, $page)
        }catch{
            Write-Host "Error with page #$page" + $_.Exception.Message
            pause
            $document.Close()
            exit
        }
    }
    Write-Host "Splitting PDF has finished. PDF page labels were used to name the files."
    Write-Host "Extracted files are located in $destFolder"
  }
