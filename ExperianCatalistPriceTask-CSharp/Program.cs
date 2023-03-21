using ExperianCatalistPriceTask_CSharp.Utility;
using ExperianCatalistPriceTask_CSharp.Excel_Manipulation;
using OfficeOpenXml;

// You must declare that you are using EPPlus with a noncommercial license. 
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

bool ErrorDetected;
// First, the program will try to find + download the spreadsheet from Experian.
// If it finds it and successfully downloads it, it enters this if statement:
if (await GetEmailAttachment.DownloadExperianCatalistFileAsync()) // DownloadEmailSpreadsheet returns 'true' if an Email is found. 
{
    // It will then try to merge the data. This function will return 'false' if it fails. 
    if (!Excel.MergeDataToSpreadsheet())
    {
        // If an error has been detected:
        ErrorDetected = true;
        GetEmailAttachment.SendEmail(ErrorDetected);
    }
    // No error detected, program has completed successfully and an Email will be sent. 
    else { ErrorDetected = false; GetEmailAttachment.SendEmail(ErrorDetected); }
}
// Else, if no Email has been found:
else
{
    // Send email.
    ErrorDetected = true;
    GetEmailAttachment.SendEmail(ErrorDetected); // The SendEmail function uses a boolean to determine if an error was detected. 
}