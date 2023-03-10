using ExperianCatalistPriceTask_CSharp.Utility;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExperianCatalistPriceTask_CSharp.Excel_Manipulation
{
    public class Excel
    {   
        // This is the where the source (the Excel sheet received from Experian) will be saved. Current working directory (CWD) is the folder that the .exe is ran from. 
        private static readonly string sourceFilepath = Directory.GetCurrentDirectory() + "\\Experian Catalist Price Averages.xlsx";

        // This is the destination/target that the data will be copied to: (WILL NEED TO BE CHANGED TO MATCH HOST PC'S DIRECTORY)
        private static readonly string targetFilepath = "C:\\Portland\\Fuel Trading Company\\Portland - Portland\\Prices\\Pump Prices vs Platts.xlsx";
        // Alternative path: "C:\\Portland\\OneDrive - Fuel Trading Company\\Portland\\Prices\\Pump Prices vs Platts.xlsx"

        // MILES'S TEST PATH!!
        //private static readonly string targetFilepath = "C:\\Users\\MilesVellozzo\\Desktop\\Pump Prices vs Platts.xlsx";

        // Open the workbooks and the specified spreadsheets:
        private static readonly ExcelPackage source = new(sourceFilepath);
        private static readonly ExcelPackage target = new(targetFilepath);
        private static readonly ExcelWorksheet sourceWs = source.Workbook.Worksheets[0];
        private static readonly ExcelWorksheet targetWs = target.Workbook.Worksheets["Imports"];

        // This is the main function responsible for copying data from the source to the target:
        public static bool MergeDataToSpreadsheet()
        {   
            // Begin by preparing the destination/target worksheet for merging:
            DeleteEmptyRows();
            StringBuilderPlusConsole.EmailBodyBuilderSBOnly("<p>Target Destination: Portland > Prices >  <b>Pump Prices vs Platts.xlsx</b></p> <hr>");
            // Find out how many rows of data exist in the target spreadsheet:
            int targetRowCount = targetWs.Dimension.End.Row;
            // Create a counter for Email logging:
            int i = 1;
            // Check if the source data is already in the target spreadsheet:
            if (!DoesDateExistInTarget(sourceWs.Cells["A2"].Text, targetRowCount))
            {
                // If it isn't, then copy over the cell:
                ExcelRange first_row = sourceWs.Cells["A2:E2"];
                ExcelRange targetRange = targetWs.Cells["A" + Convert.ToString(targetRowCount + i) + ":E" + Convert.ToString(targetRowCount + i)];
                targetRange.Value = first_row.Value;
                // Format the first cell so that it matches the destination formatting (Excel Date formatting):
                FormatDateCells("A" + Convert.ToString(targetRowCount + i));
                i++;
            }
            if (!DoesDateExistInTarget(sourceWs.Cells["A3"].Text, targetRowCount))
            {
                ExcelRange second_row = sourceWs.Cells["A3:E3"];
                ExcelRange targetRange = targetWs.Cells["A" + Convert.ToString(targetRowCount + i) + ":E" + Convert.ToString(targetRowCount + i)];
                targetRange.Value = second_row.Value;
                FormatDateCells("A" + Convert.ToString(targetRowCount + i));
                i++;
            }
            // Begin building the success Email StringBuilder:
            StringBuilderPlusConsole.EmailBodyBuilder("The program has ran <b>sucessfully</b> with the following outcome:");
            switch (i)
            {
                case 1: // No prices added - prices already exist in target spreadsheet.
                    StringBuilderPlusConsole.EmailBodyBuilder("No new prices have been added as <b>both exist in the target spreadsheet.</b>");
                    break;
                case 2: // One price added.
                    StringBuilderPlusConsole.EmailBodyBuilder("<b>[1/2]</b> new prices have been added to the target spreadsheet. One price <b>already existed</b> in the target spreadsheet.");
                    break;
                case 3: // Two prices added.
                    StringBuilderPlusConsole.EmailBodyBuilder("<b>[2/2]</b> new prices have been added to the target spreadsheet.");
                    break;
            }
            try
            {   
                // This will only fail if the target Excel sheet is currently open:
                target.Save();
            }
            catch
            {
                // If it is open, then create the error Email StringBuilder:
                StringBuilderPlusConsole.ErrorEmailBodyBuilderSBOnly("<p>Target Destination: Portland > Prices > <b>Pump Prices vs Platts.xlsx</b></p> <hr>");
                StringBuilderPlusConsole.ErrorEmailBodyBuilder("The program has encountered a <b>critical error</b> and could not complete its task:");
                StringBuilderPlusConsole.ErrorEmailBodyBuilder("The <b>target spreadsheet is open</b>, and thus cannot be manipulated by this program.");
                StringBuilderPlusConsole.ErrorEmailBodyBuilder("This program will need to be <b>ran again manually</b> once the <i>offending user</i> vacates the spreadsheet.");
                return false;
            }
            return true;
        }
        private static void FormatDateCells(string cell)
        {
            DateOnly cellDate = DateOnly.ParseExact(targetWs.Cells[cell].Text, "dd/MM/yyyy");
            string excelDateFunction = $"=DATE({cellDate.Year}, {cellDate.Month}, {cellDate.Day})";
            targetWs.Cells[cell].Formula = excelDateFunction;
            targetWs.Cells[cell].Style.Numberformat.Format = "dd/MM/yyyy";
            Console.WriteLine("Formatted cell " + cell + " in target spreadsheet to date format.");
        }
        private static bool DoesDateExistInTarget(string date, int targetRowCount)
        {
            DateOnly lastDateInTarget = DateOnly.ParseExact(targetWs.Cells["A" + targetRowCount].Text, "dd/MM/yyyy");
            DateOnly dateToCheck = DateOnly.ParseExact(date, "dd/MM/yyyy");
            if (dateToCheck <= lastDateInTarget)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        // This function deletes empty rows that are found in the target worksheet.
        // Occasionally, when a user manually manipulates an Excel sheet they leave rows behind that appear empty - but they have null values. 
        // This function just ensures that there are no occurances, so that merged data can be appended correctly. 
        private static void DeleteEmptyRows()
        {
            int amountOfEmptyRows = 0;
            // Loop through each row of the worksheet:
            for (int i = targetWs.Dimension.Start.Row; i <= targetWs.Dimension.End.Row; i++ )
            {   
                // Check if the first cell of the row is empty. If it is, this indicates it is a null row. 
                if (targetWs.Cells["A" + i].Value == null)
                {
                    amountOfEmptyRows++;
                    targetWs.DeleteRow(i);
                    // Decrement the row counter so that the next iteration does not skip a row. 
                    i--;
                }
            }
            if (amountOfEmptyRows>0)
            {
                Console.WriteLine(amountOfEmptyRows + " empty rows have been found. These will be deleted.");
            }
            else { Console.WriteLine("No empty rows have been found. Proceeding with the merge..."); }
        }
    }
}
