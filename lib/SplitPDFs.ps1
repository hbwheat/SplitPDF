#LIVE ON JESTER Version 0.7

#Change Log
<#
    2014.4.20 - Added "$pdfname-Extracted" to output file name to avoid naming conflicts
    2017.4.19 - Added exit if no page labels exist 
    
#>
#Description
<#
    This script uses the iTextSharp c# library to split a pdf into seperate pages and rename them
    based on the pdf lables (page names), if existing. This library powers the copy of each page. The library can also
    split mult pages if necessary. 

    Eventually the script will take into account Bookmarks (outlines) and identify any page numbers referenced. The script 
    will then make a best guess at the page names based of user input reference of page number scheme. 

    I used man References search the tag REF to find all links. 
    Source for library here: https://sourceforge.net/projects/itextsharp/
    Index for library: http://afterlogic.com/mailbee-net/docs-itextsharp/Index.html
#>


<#
REF: https://www.codeproject.com/Articles/559380/Splitting-and-Merging-PDF-Files-in-Csharp-Using-iT
#> #REF

Write-Host "Launching Split PDFs Utility...."

#TODO: need to use $PSScriptroot to get root of current dir
#############################################################################
##This will need to reference the root folder of the location of the script##
##The .dll and .xml for the itextsharp library will need to be in ~/lib/   ##
#############################################################################
$dir = "\\Jester\Software\Miscellaneous\Scripts\SplitPDFs"
#import dll
Add-Type -Path $dir\lib\itextsharp.dll

<#
REF: https://www.codeproject.com/Articles/559380/Splitting-and-Merging-PDF-Files-in-Csharp-Using-iT
#> #REF: for C#
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
    $OpenFileDiag.Filter = "PDF (*.pdf)| *.pdf"
    $OpenFileDiag.ShowDialog() | Out-Null
    $OpenFileDiag.FileName
}
#window open initial location
$sourcePDFPath = get-Filename -initialDir $env:USERPROFILE\Desktop

#test document opening
try{
    $document = New-Object iTextSharp.text.pdf.PdfReader -ArgumentList $sourcePDFPath -ErrorAction Stop
    $docTotalPages = $document.NumberOfPages
    $arrayofLabels = [iTextSharp.text.pdf.PdfPageLabels]::GetPageLabels($document)
    #$document.Close()
}catch{
    Write-Host "Error importing PDF file"
    pause
    exit
    #TODO: add error correction opportunity
}

#drive the script and the save off file names
#do if when labels are for every page, otherwise do else
if($arrayofLabels.Length -ne $docTotalPages){
    if($arrayofLabels.length -eq 0){
        Write-Host "This PDF has no page labels! Exiting now"
        pause
        exit
    }
    #need to adjust page numbers. Use chapters and incrememnt for each page in the chapter
    #not enough label names in PDF
    #TODO: Add what if no labels for each page
    $page = 0
    while($page -lt $docTotalPages){
        if($arrayofLabels[$page] -eq $arrayofLabels[$page + 1]){
            #check next label for the chapter 
        }
    }

}else{ #files are good to go to be named and saved
    #Set path labels
    $folderName = [regex]::Split($sourcePDFPath,'.pdf$') #remove .pdf tag
    $folderName = %{$split = [regex]::Match($folderName,'[\w\.\ ]*$'); $split[0].Value} #select the name up to the \ in the path right to left
    $folderName = [regex]::Split($folderName,'[\s]$') -join "-Extracted" #remove white space at end (for seme reason left) probably regex on line 219
    #create save to folder in root where sourcePDF is located called "$pdfname-Extracted"
    #TODO: instead of regex for name of PDF use $document.
    $destFolder = [regex]::Split($sourcePDFPath,'[\w\d\s]*.pdf$') -join $folderName 
    New-Item $destFolder -type Directory
    $page = 0
    #loop through all papes
    while($page -lt $docTotalPages){
        Write-Progress -Activity "Spliting PDF... ['ctrl+c to' force stop]" -PercentComplete ($page / $docTotalPages * 100) -CurrentOperation "Page: $page of $docTotalPages."
        #set save to dest
        #Regex to remove disallowed characters from string of path
        $pdfName = [regex]::Replace($arrayofLabels[$page], '[^\w\-\.\ ]', '')
        $savetoPath = $destFolder + "\" + $pdfName + ".pdf"
        $newPDF = [iTextTools.PdfExtractorUtility]::new()
        $page++ #count here so above path is accurate in getting label, and below is accurate on what page
        try{
            $newPDF.ExtractPage($sourcePDFPath, $savetoPath, $page)
        }catch{
            Write-Host "Error with page #$page." + $_.Exception.Message
            pause
            $document.Close()
            exit
        }
    }
}
