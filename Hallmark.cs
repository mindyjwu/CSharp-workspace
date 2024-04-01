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
            int searchColumn = 3; // Column C
            string searchTerm = "TOTALS";
            int searchStartRow = 1; // Assuming you want to start from the first row
            int totalsRow, Col_UnitsReturned;
            int titleOffset = 0; // Assuming TitleOffset is defined previously
            
            Range columnC = xlWorksheet.Columns[searchColumn];
            Range result = columnC.Find(
                What: searchTerm,
                After: xlWorksheet.Cells[searchStartRow, searchColumn],
                LookIn: XlFindLookIn.xlValues,
                LookAt: XlLookAt.xlWhole,
                SearchOrder: XlSearchOrder.xlByRows,
                SearchDirection: XlSearchDirection.xlNext,
                MatchCase: false
            );
            
            if (result != null)
            {
                // Found the row with "TOTALS"
                totalsRow = result.Row;
                // Hide Units Returned if none displayed.
                if ((double)(xlWorksheet.Cells[totalsRow, Col_UnitsReturned] as Range).Value2 == 0)
                {
                    (xlWorksheet.Columns[Col_UnitsReturned] as Range).EntireColumn.Hidden = true;
                    titleOffset += 1;
                }
}

        
              
        }
    
        private void HideColumnsWithNoDisplay()
        {
            // Logic to hide "Units Returned", "Per Item Rate", "Reserves Taken", and "Reserves Liquidated" if none displayed.
        }
    
        private void HideUnconditionallyHiddenColumns()
        {
            // Logic to unconditionally hide specific columns like "HALS-49 Contributor Share / Contributor Royalty Earned".
        }
    
        private void HideNetRoyaltiesEarnedColumnIfSame()
        {
            // Logic to hide "Net Royalties Earned" column if all values are the same for "Royalties Earned" column.
        }
    
        private void HideLogoBasedOnCompanyUDF()
        {
            // Logic to hide the logo depending on the Company UDF value.
        }
    
        public void CopyAndPasteFunction()
        {
            MoveMainStatementTitle();
            MoveRecoupmentGroup();
        }
    
        private void MoveMainStatementTitle()
        {
            // Logic to move the main statement title, find unhidden columns, copy and paste the title, and clear the original title.
        }
    
        private void MoveRecoupmentGroup()
        {
            // Logic to move Recoupment Group from column B into column C by iterating through all rows and searching for "RG@@" text.
        }
    }
}
                
