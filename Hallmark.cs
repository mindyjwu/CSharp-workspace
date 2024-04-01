using Aspose.Cells;
using Aspose.Cells.Properties;
using Aspose.Cells.Rendering;
using SpreadsheetGear;
using SpreadsheetGear.Themes;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
namespace PostExecuteProcessing
{
    public class PostExecuteProcessing
    {
        /// <summary>
        /// Opens an Excel file and performs any Post Execute Processing necessary. Saves and closes the file.
        /// </summary>
        
        public void Execute(string workbookPath)
        {
            /// HAL-500 Initial config (Mindy & Ajay, April-01-2024)
            IWorkbook workbook = Factory.GetWorkbookSet().Workbooks.Open(workbookPath);
            try
            {
                IWorksheet worksheet = workbook.Worksheets["Sheet1"];
                IRange cells = worksheet.Cells;
                // Unprotect the sheet if needed
                bool sheetProtected = worksheet.ProtectContents;
                if (sheetProtected)
                {
                    worksheet.Unprotect("");
                }
                // Note method refactored later for re-use, the original one matches assignment examples
                //CopyContractInfo(worksheet);

                //ClearOriginalTitle(worksheet);
                //InsertRecoupment()
                
                worksheet.Cells["B1"].EntireColumn.Hidden = true;
                worksheet.Cells["J1"].EntireColumn.Hidden = true;

                // Protect the sheet again, if it was protected
                if (sheetProtected)
                {
                    worksheet.Protect("");
                }
                // Save workbook
                workbook.Save();
            }
            finally
            {
                // Always close, even if there is an exception
                workbook.Close();
            }

            using (Workbook asposeWorkbook = new Workbook(workbookPath))
            {
                Worksheet worksheet = asposeWorkbook.Worksheets["Sheet1"];
                asposeWorkbook.Save(workbookPath);
            }
        }

        // Common Wrapper for Searching 
        private IRange FindNextMatch(IRange cells, string columnStr, int startRow, string term, LookAt lookAt)
        {
            int column = CellsHelper.ColumnNameToIndex(columnStr);
            return cells[1, column].EntireColumn.Find(
                    term, cells[startRow, column], FindLookIn.Values, lookAt,
                    SearchOrder.ByRows, SearchDirection.Next, false);
        }
        private IRange FindNextPartialMatch(IRange cells, string columnStr, int startRow, string term)
        {
            return FindNextMatch(cells, columnStr, startRow, term, LookAt.Part);
        }
        private IRange FindNextExactMatch(IRange cells, string columnStr, int startRow, string term)
        {
            return FindNextMatch(cells, columnStr, startRow, term, LookAt.Whole);
        }
        // --------------------------------- Hide Function ------------------------- 
                //HideTotal(worksheet);
                //HidePerItem(worksheet);
                //HideReserveTaken(worksheet);
                //HideReserveLiquidated(worksheet);
                //HideNetRoyaltiesEarned(worksheet);
                //HideLogo(worksheet);
        
        public void HideColumnsBasedConditions()
        {
            HideIfFunction();
            HideColumnsNoDisplay();
            HideHiddenColumns();
            HideNetRoyaltiesEarned();
            HideLogoUDF();
        }
                
          private void HideIfFunction()
        {
            // find the row with "TOTALS" and hide specific columns based on conditions.
            // iterating over the cells and checking values
            int searchStartRow = 1; // Assuming you want to start from the first row
            int totalsRow, Col_UnitsReturned;
            int titleOffset = 0; // Assuming TitleOffset is defined previously
              
            while (searchStartRow <= worksheet.UsedRange.Range.RowCount)
            {
                IRange result = FindNextExactMatch(worksheet.Cells, "C", searchStartRow, "Total")
            }
            if (result != null)
            {
                // Found the row with "TOTALS"
                totalsRow = result.Row;
                // Hide Units Returned if none displayed.
                if ((double)(worksheet.Cells[totalsRow, Col_UnitsReturned] as Range).Value2 == 0)
                {
                    (xlWorksheet.Columns[Col_UnitsReturned] as Range).EntireColumn.Hidden = true;
                    titleOffset += 1;
                }
            }        
        }
    
        private void HideColumnsWithNoDisplay()
        {
            // hide "Units Returned", "Per Item Rate", "Reserves Taken", and "Reserves Liquidated" if none displayed
            Range totalsRowRange = FindNextExactMatch(worksheet.Columns[1], "C", 1, searchTerm);
            if (totalsRowRange != null)
            {
                int totalsRow = totalsRowRange.Row;
                HideColumnIfZero(xlWorksheet, totalsRow, "B", ref titleOffset); // Col_UnitsReturned
                HideColumnIfZero(xlWorksheet, totalsRow, "C", ref titleOffset); // Col_PerItemRate
                HideColumnIfZero(xlWorksheet, totalsRow, "D", ref titleOffset); // Col_ReservesTaken
                HideColumnIfZero(xlWorksheet, totalsRow, "E", ref titleOffset); // Col_ReservesLiquidated
            }
        }
        
        private void HideColumnIfZero(Worksheet worksheet, int row, string columnStr, ref int titleOffset)
        {
            int column = CellsHelper.ColumnNameToIndex(columnStr);
            if ((xlWorksheet.Cells[row, column] as Range).Value2 == 0)
            {
                worksheet.Columns[column].EntireColumn.Hidden = true;
                titleOffset--; // Adjust based on your logic, could be increment or decrement
            }
            else if (columnStr == "C") // Assuming "C" is the "Per Item Rate" column
            {
                // Grandtotal is not required for "Per Item Rate" column
                (worksheet.Cells[row, column] as Range).Value2 = "";
            }
        }
        
        private void HideUnconditionallyHiddenColumns()
        {
            // unconditionally hide specific columns like "HALS-49 Contributor Share / Contributor Royalty Earned"
            int Col_ContributorShare = CellsHelper.ColumnNameToIndex("S");
            worksheet.Columns[Col_ContributorShare].EntireColumn.Hidden = true;

            // Adjust title offset
            titleOffset -= 2;
        
            // Clear the 'Grandtotal' for Contributor Share column
            worksheet.Cells[TotalsRow, Col_ContributorShare].Value = "";
            
        }
    
        private void HideNetRoyaltiesEarnedColumnIfSame()
        {
            // hide "Net Royalties Earned" column if all values are the same for "Royalties Earned" column
            int sameFlag = 0;
            int firstRow = 9; // Starting row for data
            int targetCol = 18; // Column R
            int hideCol = 23; // Column W
            int lastRow = worksheet.Cells[worksheet.Rows.Count, targetCol].End(XlDirection.xlUp).Row;
        
            for (int counter = firstRow; counter <= lastRow; counter++)
            {
                if (worksheet.Cells[counter, targetCol].Value2.ToString() != worksheet.Cells[counter, hideCol].Value2.ToString())
                {
                    sameFlag = 1;
                    break;
                }
            }
        }
        
        private void HideLogoBasedOnCompanyUDF()
        {
            // hide the logo depending on the Company UDF value
            int searchRow = 1; // row where the company name is located
            int searchCol = 53; // column BA
            // Get the UDF value from BA, if not set default to alliantLogo
            string udfValue = worksheet.Cells["BA"].Text;
            
            switch(udfValue)
            {
                default:
                    worksheet.Shapes["Picture 1"].Visible = true;
                    worksheet.Shapes["Picture 2"].Visible = false;
                    break;
                case "Hallmark":
                    worksheet.Shapes["Picture 1"].Visible = true;
                    worksheet.Shapes["Picture 2"].Visible = false;
                    break;

                case "DaySpring":
                    worksheet.Shapes["Picture 1"].Visible = false;
                    worksheet.Shapes["Picture 2"].Visible = true;
                    break;
            } 
        }
    
        public void CopyAndPasteFunction()
        {
            MoveMainStatementTitle();
            MoveRecoupmentGroup();
        }
    
        private void MoveMainStatementTitle()
        {
            // move the main statement title, find unhidden columns, copy and paste the title, and clear the original title.
        }
    
        private void MoveRecoupmentGroup()
        {
            // move Recoupment Group from column B into column C by iterating through all rows and searching for "RG@@" text.
        }
    }
}
                
