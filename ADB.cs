using Aspose.Cells;
using Aspose.Cells.Properties;
using Aspose.Cells.Rendering;
using SpreadsheetGear;
using SpreadsheetGear.Themes;
using System;
using System.Drawing;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Cells.Drawing;
using SpreadsheetGear.Charts;

namespace PostExecuteProcessing
{
    /// <summary>
    /// Performs any Post Execute Processing on an Excel file
    /// 
    /// Notes:
    /// The namespace, class name and Execute method name are all used by the Statement Service and cannot be changed.
    /// Additional external libraries will likely require a change to the Statement Service to include them.
    /// 
    /// You should only use System.Drawing to reference Color when working with Aspose.Cells. See AD-18503 for more information on how
    /// it can potentially cause issues.
    /// </summary>
    public class PostExecuteProcessing
    {
        /// <summary>
        /// Opens an Excel file and performs any Post Execute Processing necessary. Saves and closes the file.
        /// </summary>
        public void Execute(string workbookPath)
        {
            /// Open in SpreadsheetGear
            ///ADB - 584 ASO Config (Mindy Wu, June-2024)
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
                // Note method refactored later for re-use, the original one matches 
                HideColumns((workbook);
                SKUValues(workbook);
                HideZeroRows(workbook);
                HideBlankRow(workbook);
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

        }


        // Find Function that return IRange of cells based on the column String and search term  
        private IRange FindNextMatch(IRange cells, string columnStr, int startRow, string term, LookAt lookAt)
        {
            int column = CellsHelper.ColumnNameToIndex(columnStr);
            return cells[1, column].EntireColumn.Find(
                    term, cells[startRow, column], FindLookIn.Values, lookAt,
                    SearchOrder.ByRows, SearchDirection.Next, false);
        }
        private IRange FindNextExactMatch(IRange cells, string columnStr, int startRow, string term)
        {
            return FindNextMatch(cells, columnStr, startRow, term, LookAt.Whole);
        }

        private void HideColumns(IWorkbook workbook)
        {
            // initialize variables
            IRange result;
            int findRow, lastRow, findLastRow, searchCol, columnCount, searchStartRow;
            string searchTerm;
            IWorksheet worksheet = workbook.Worksheets["Sheet1"];
            //int colA = CellsHelper.ColumnNameToIndex("A");
            string colA = "A";
            columnCount = 0;
            searchCol = 1;
            searchStartRow = 0;

            /** With worksheet 1 **/
            // set result as nothing
            result = null;
            // result = sheet1, Z, 1
            searchTerm = worksheet.Cells[1, 26].Value.ToString();
            // get target row data
            //find the search term by rows
            result = FindNextExactMatch(worksheet.Cells, colA, searchStartRow, searchTerm);
            
            if (result != null)
            {
                // if NOT result is nothing then, set findRow as result.row
                findRow = result.Row;
            }
            else
            {
                // exist if the search term is not found
                return;
            }

            // get target row data 
            // find searchterm = “End Here 2” by rows
            searchTerm = "End Here 2";
            // set search column to col P
            string colP = "P";
            searchCol = 16;
            // Hide search column P
            worksheet.Cells["P:P"].EntireColumn.Hidden = false;
            result = FindNextExactMatch(worksheet.Cells, colP, searchStartRow, searchTerm);
            worksheet.Cells["P:P"].EntireColumn.Hidden = true;
            // if NOT result is nothing then, set findLastRow = result.row
            if (result != null)
            {
                findLastRow = result.Row;
                // update last row = findlastRow - findRow + 1
                lastRow = findLastRow - findRow + 1;
            }
            else
            {
                return;
            }

            /** With worksheet 2 **/
            IWorksheet worksheet2 = workbook.Worksheets["Sheet2"];
            // copy format from ‘sheet1’ to ‘sheet2’
            //IRange sourceRange = worksheet.Cells["A" + findRow + ":L" + findLastRow];
            sourceRange.Copy();
            //worksheet.IRange("A" + findRow + ":L" + findLastRow).Copy();
            worksheet2.IRange("A1:L" + lastRow).PasteSpecial(XlPasteType.xlPasteAll);
            // copy value ‘sheet1’ to ‘sheet2’
            worksheet.IRange("A" + findRow + ":L" + findLastRow).Copy();
            worksheet2.IRange("A1:L" + lastRow).PasteSpecial(XlPasteType.xlPasteValues);

            // if column AB of ‘sheet1’ is not ‘Yes’, then delete corresponding column in ‘sheet2’
            // increment columnCount 
            //update column count
            // UDF Was Removed, and Value is always "No"
            
            for (int i = 1; i <= 8; i++)
            {
                if (worksheet.Cells[i, 28].Value != "Yes")
                {
                    worksheet2.Columns[i].EntireColumn.Delete();
                    columnCount++;
                }
            }
            worksheet.Cells[2,24].Value = columnCount;
            // copy and paste data form ‘sheet2’ back to range in ‘sheet1’
            worksheet2.IRange("A1:K" + lastRow).Copy();
            worksheet2.IRange("A" + findRow + ":K" + findLastRow).PasteSpecial(XlPasteType.xlPasteAllUsingSourceTheme);

        }

        private void SKUValues(IWorkbook workbook)
        {
            // initialize variables
            IRange result;
            int findRow, startRow, lastRow, counter, findLastRow, searchCol;
            string searchTerm, previousString, findString;

            //int colA = CellsHelper.ColumnNameToIndex("A");
            string colA = "A";


            /** with worksheet 3 **/
            IWorksheet worksheet3 = workbook.Worksheets["Sheet3"];
            // find last row with data in column A
            lastRow = worksheet3.Cells[worksheet3.Cells.Rows.Count - 1, 0].End(SpreadsheetGear.Enums.Direction.Up).Row;
            // initialize PreviousString with value in first cell of column A
            previousString = worksheet3.Cells[0, 0].Text;
            // loop through rows from 2 to lastRow to hide (deletion) rows with duplicate values
            for (int row = 1; row <= lastRow; row++)
            {
                string currentString = worksheet3.Cells[row, 0].Text;
                if (currentString == previousString)
                {
                    worksheet3.Cells[row, 0].EntireRow.Hidden = true;
                }
                else
                {
                    previousString = currentString;
                }
            }

            // loop through rows from LastRow to 2 in reverse order to delete the hidden rows
            for (int row = lastRow; row >= 1; row--)
            {
                if (worksheet3.Cells[row, 0].EntireRow.Hidden)
                {
                    worksheet3.Cells[row, 0].EntireRow.Delete();
                }
            }


            /** with worksheet 1**/
            IWorksheet worksheet = workbook.Worksheets["Sheet1"];
            // search ‘Start Here’ in col M
            searchTerm = "Start Here";
            result = worksheet.Cells.Find(searchTerm, LookIn: FindLookIn.Values, LookAt: FindLookAt.Whole, SearchOrder: SearchOrder.ByRows, SearchDirection: SearchDirection.Next);
            if (result != null)
            {
                // set startRow to row number after found row
                startRow = result.Row + 1;
            }
            else
            {
                // Exit if 'Start Here' is not found
                return;
            }

            // search ‘End Here’ in col M
            searchTerm = "End Here";
            result = worksheet.Cells.Find(searchTerm, LookIn: FindLookIn.Values, LookAt: FindLookAt.Whole, SearchOrder: SearchOrder.ByRows, SearchDirection: SearchDirection.Next);
            if (result != null)
            {
                // set findLastRow to row number before found row 
                findLastRow = result.Row - 1;
            }
            else
            {
                // Exit if 'End Here' is not found
                return;
            }
            // loop through rows from startRow to findLastRow to find and update values based on search term in ‘sheet3’
            for (int row = startRow; row <= findLastRow; row++)
            {
                findString = worksheet.Cells[row, 12].Text; // Column M is the 13th column (index 12)
                result = worksheet3.Cells.Find(findString, LookIn: FindLookIn.Values, LookAt: FindLookAt.Whole, SearchOrder: SearchOrder.ByRows, SearchDirection: SearchDirection.Next);
                if (result != null)
                {
                    worksheet.Cells[row, 12].Value = result.Value;
                }
            }
        }

        private void HideZeroRows(IWorkbook workbook)
        {
            // initialize variables
            IRange result;
            int hiddenCol, netSalesCol, royaltiesDueCol, startRow, endRow, counter;
            string searchTerm;

            /** with worksheet1 **/
            IWorksheet worksheet = workbook.Worksheets["Sheet1"];
            // check specific cells in column AB (28th column)
            hiddenCol = 0;
            for (int row = 0; row < worksheet.Cells.Rows.Count; row++)
            {
                if (worksheet.Cells[row, 27].Text != "Yes")
                {
                    // increment HiddenCol by one if the cell value is not ‘Yes’
                    hiddenCol++;
                }
            }

            // set initial column indices for Net Sales and Royalties Due
            netSalesCol = 10;    //column J
            string colJ = "J";
            royaltiesDueCol = 12;   //column L
            string colL = "L";
            // adjust netsales based on hidden columns
            netSalesCol = netSalesCol - hiddenCol;
            int searchCol = 1;
            int searchStartRow = 0;

            // define the search term (ROYALTY DETAILS) to locate the starting row
            searchTerm = "ROYALTY DETAILS";
            result = FindNextExactMatch(worksheet.Cells, colJ, searchStartRow, searchTerm);
            // if search term is found, then set the start row
            if (result != null)
            {
                startRow = result.Row + 3;
            }
            else
            {
                // else, exist the subroutine if not found
                return;
            }

            // adjust the column indices based on ‘net sales’ in header row
            // if cell in row above StartRow and NetSalesCol contains text ‘Net Sales’, then adjust RoyDueCol - HiddenCol to account hidden columns
            bool netSalesFound = false;
            for (int col = 0; col < worksheet.Cells.Columns.Count; col++)
            {
                if (worksheet.Cells[startRow - 1, col].Text == "Net Sales")
                {
                    netSalesCol = col;
                    royaltiesDueCol -= hiddenCol;
                    netSalesFound = true;
                    break;
                }
            }
            // Check again if cell in row above startRow and adjusted netSalesCol contains 'Net Sales'
            if (!netSalesFound || worksheet.Cells[startRow - 1, netSalesCol].Text != "Net Sales")
            {
                // If 'Net Sales' not found in adjusted column, exit the subroutine
                return;
            }

            // Calculate the end row
            endRow = worksheet.Cells[worksheet.Cells.Rows.Count - 1, netSalesCol].End(SpreadsheetGear.Enums.Direction.Up).Row;

            // else, if the ‘Net Sales’ not found in adjusted column, exist the subroutines 
            // Loop through rows from startRow to endRow
            for (counter = startRow; counter <= endRow; counter++)
            {
                if (worksheet.Cells[counter, netSalesCol].Value != null && worksheet.Cells[counter, royaltiesDueCol].Value != null &&
                    Convert.ToDouble(worksheet.Cells[counter, netSalesCol].Value) == 0 &&
                    Convert.ToDouble(worksheet.Cells[counter, royaltiesDueCol].Value) == 0)
                {
                    worksheet.Cells[counter, 0].EntireRow.Hidden = true;
                }
            }

        }

        private void HideBlankRow(IWorkbook workbook)
        {
            // Initialize variables
            IRange result;
            String searchTerm;
            int searchCol = 1; // column A
            String colA = "A";
            int findRow = -1;
            int findLastRow = -1;
            int searchStartRow = 0;

            /** with worksheet 1 **/
            IWorksheet worksheet = workbook.Worksheets["Sheet1"];
            // get target row data
            // set result as nothing
            result = null;
            // find the search term "Beg. Minimum Balance" by rows
            searchTerm = "Beg. Minimum Balance";
            result = FindNextExactMatch(worksheet.Cells, colA, searchStartRow, searchTerm);
            // if NOT result is nothing then, set findRow as result.row
            if (result != null)
            {
                findRow = result.Row;
            }
            else
            {
                // Exit if the search term is not found
                return;
            }


            // get target row data
            // find search term "Beg. Maximum Balance" by rows
            searchTerm = "Beg. Maximum Balance";
            FindNextExactMatch(worksheet.Cells, colA, searchStartRow, searchTerm);
            // if NOT result is nothing then, set findRow as result.row
            if (result != null)
            {
                findRow = result.Row;
            }
            else
            {
                // Exit if the search term is not found
                return;
            }

            // get target row data
            // find search term "End Here 2" by rows
            String searchTerm = "End Here 2";
            FindNextExactMatch(worksheet.Cells, colA, searchStartRow, searchTerm);
            // if NOT result is nothing then, set findLastRow as result.row
            if (result != null)
            {
                findLastRow = result.Row;
                // update last row = findLastRow - findRow + 1
                int lastRow = findLastRow - findRow + 1;
            }

            // loop through rows between findRow and findLastRow
            // if row is blank, hide the row
            for (int row = findRow; row <= findLastRow; row++)
            {
                // Check if the row is blank
                bool isBlank = true;
                for (int col = 0; col < worksheet.Cells.Columns.Count; col++)
                {
                    if (!string.IsNullOrEmpty(worksheet.Cells[row, col].Text))
                    {
                        isBlank = false;
                        break;
                    }
                }

                // If the row is blank, hide the row
                if (isBlank)
                {
                    worksheet.Cells[row, 0].EntireRow.Hidden = true;
                }
            }
            // end loop

        }

        private void MoveLeftColumns(IWorkbook workbook)
        {
        }
        private void DisplayDateRange(IWorkbook workbook)
        {
        }
    }
}
