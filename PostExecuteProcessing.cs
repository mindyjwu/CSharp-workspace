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

namespace PostExecuteProcessing
{
    /// <summary>
    /// Performs any Post Execute Processing on an Excel file
    /// 
    /// Notes:
    /// The namespace, class name and Execute method name are all used by the Statement Service and cannot be changed.
    /// Additional external libraries will likely require a change to the Statement Service to include them.
    /// </summary>
    public class PostExecuteProcessing
    {
        /// <summary>
        /// Opens an Excel file and performs any Post Execute Processing necessary. Saves and closes the file.
        /// </summary>
        public void Execute(string workbookPath)
        {
            /// Section Name: A01 Assignment 1 (with C# Assignment)
            /// 
            /// RI-A002 RI-123 Initial config (Jude Saldanha, Dec-06-2023)
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
                //CreateBorderInfo(worksheet);
                CreateBorderInfoNew(worksheet);
                CopyContractInfo(worksheet);
                //HideZeroReturns(worksheet);
                HideZeroReturnsNew(worksheet);
                CreateBorderTotal(worksheet);
                InsertPart12(worksheet);
                //GetLogo(worksheet);
                GetLogoNew(worksheet);
                CreateBorderAroundPart(worksheet, "Part 10");
                CreateBorderAroundPart(worksheet, "Part 11");
                // Hide Columns B and J
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

                PageBreakAfterPart9(worksheet);

                //var pageBreaks = worksheet.GetPrintingPageBreaks(new ImageOrPrintOptions());
                //Console.Out.WriteLine("Page Break Count" + worksheet.GetPrintingPageBreaks(new ImageOrPrintOptions()).Count());

                // Add Page Number
                worksheet.PageSetup.SetFooter(1, "Page &P");

                asposeWorkbook.Save(workbookPath);
            }
        }

        /// RI-A002 RI-123 Create the border for the info part of the statement (Jude Saldanha, Dec-06-2023) <summary>
        /// </summary>
        /// <param name="worksheet"></param>
        private void CreateBorderInfo(IWorksheet worksheet)
        {
            // Cells F2 to I4
            var borderRange = RangeToRange(worksheet.Cells["F2"], worksheet.Cells["I4"]);

            // Color index uses a zero based index
            borderRange.Borders[BordersIndex.EdgeBottom].ColorIndex = 0;
            borderRange.Borders[BordersIndex.EdgeBottom].Weight = BorderWeight.Medium;
            borderRange.Borders[BordersIndex.EdgeTop].ColorIndex = 0;
            borderRange.Borders[BordersIndex.EdgeTop].Weight = BorderWeight.Medium;
            borderRange.Borders[BordersIndex.EdgeLeft].ColorIndex = 0;
            borderRange.Borders[BordersIndex.EdgeLeft].Weight = BorderWeight.Medium;
            borderRange.Borders[BordersIndex.EdgeRight].ColorIndex = 0;
            borderRange.Borders[BordersIndex.EdgeRight].Weight = BorderWeight.Medium;
        }

        /// RI-A002 RI-123 Create the border for the info part of the statement (Jude Saldanha, Dec-06-2023)
        private void CreateBorderInfoNew(IWorksheet worksheet)
        {
            CreateBorderBlackMedium(worksheet, "F2", "I4");
        }

        /// RI-A002 RI-123 Common method to draw a black border (Jude Saldanha, Dec-06-2023)
        private void CreateBorderBlackMedium(IWorksheet worksheet, string startCell, string endCell)
        {
            DrawBorder(worksheet, startCell, endCell,
                new BordersIndex[] { BordersIndex.EdgeBottom, BordersIndex.EdgeTop, BordersIndex.EdgeLeft, BordersIndex.EdgeRight },
                0, // Black
                BorderWeight.Medium);
        }

        /// RI-A002 RI-123 Common method to draw a thick black border (Jude Saldanha, Dec-19-2023)
        private void CreateBorderBlackThick(IWorksheet worksheet, string startCell, string endCell)
        {
            DrawBorder(worksheet, startCell, endCell,
                new BordersIndex[] { BordersIndex.EdgeBottom, BordersIndex.EdgeTop, BordersIndex.EdgeLeft, BordersIndex.EdgeRight },
                0, // Black
                BorderWeight.Thick);
        }

        /// RI-A002 RI-123 Common Method for drawing borders (Jude Saldanha, Dec-06-2023)
        private void DrawBorder(IWorksheet worksheet, string startCell, string endCell,
            BordersIndex[] boarderIndexes, int colorIndex, BorderWeight weight)
        {
            var borderRange = RangeToRange(worksheet.Cells[startCell], worksheet.Cells[endCell]);

            // Assign the color and weight to all the input indexes
            foreach (var index in boarderIndexes)
            {
                borderRange.Borders[index].ColorIndex = colorIndex;
                borderRange.Borders[index].Weight = weight;
            }
        }

        /// RI-A002 RI-123 Copy the contract info to 2 rows after end (Jude Saldanha, Dec-06-2023)
        private void CopyContractInfo(IWorksheet worksheet)
        {
            // Find last row, add 2
            int targetRowStart = worksheet.UsedRange.Range.RowCount + 2;
            // End 2 rows later
            int targetRowEnd = targetRowStart + 2;
            string targetRange = "F" + targetRowStart + ":I" + targetRowEnd;
            worksheet.Cells["F2:I4"].Copy(worksheet.Cells[targetRange]);

            // Add thick border
            CreateBorderBlackThick(worksheet, "F" + targetRowStart, "I" + targetRowEnd);
        }

        /// RI-A002 RI-123 Hide Zero Returns in Column F (Jude Saldanha, Dec-07-2023)
        private void HideZeroReturns(IWorksheet worksheet)
        {
            // Start from row 7
            int searchStartRow = 7;
            int searchColumn = CellsHelper.ColumnNameToIndex("G");
            string searchTerm = "0.00";

            // Keep searching until we get to the end of the Used Range
            while (searchStartRow <= worksheet.UsedRange.Range.RowCount)
            {
                // Search in Column for next match
                IRange result = worksheet.Cells[1, searchColumn].EntireColumn.Find(
                        searchTerm, worksheet.Cells[searchStartRow, searchColumn], FindLookIn.Values, LookAt.Whole,
                        SearchOrder.ByRows, SearchDirection.Next, false);

                // Terminate loop if nothing found or we looped back round to the beginning
                if (result == null || result.Row < searchStartRow)
                {
                    break;
                }

                // Try and parse it into a decimal
                Decimal decValue = -1;
                bool isDecimal = Decimal.TryParse(result.Text, out decValue);

                //  Hide Row if it is a zero decimal
                if (isDecimal && decValue == 0)
                {
                    result.EntireRow.Hidden = true;
                }

                // Carry on from next row
                searchStartRow = result.Row;
            }
        }

        /// RI-A002 RI-123 Hide Zero Returns in Column F (Jude Saldanha, Dec-07-2023)
        private void HideZeroReturnsNew(IWorksheet worksheet)
        {
            // Start from row 7
            int searchStartRow = 7;

            // Keep searching until we get to the end of the Used Range
            while (searchStartRow <= worksheet.UsedRange.Range.RowCount)
            {
                // Search in Column for next match
                // Note this is looking for 0 or 0.00 just in case
                // In reality we should fix the section to always output 0.00 and do an exact match
                IRange result = FindNextPartialMatch(worksheet.Cells, "G", searchStartRow, "0");

                // Terminate loop if nothing found or we looped back round to the beginning
                if (result == null || result.Row < searchStartRow)
                {
                    break;
                }

                // Try and parse it into a decimal
                Decimal decValue = -1;
                bool isDecimal = Decimal.TryParse(result.Text, out decValue);

                //  Hide Row if it is a zero decimal
                if (isDecimal && decValue == 0)
                {
                    result.EntireRow.Hidden = true;
                }

                // Carry on from next row
                searchStartRow = result.Row;
            }
        }

        /// RI-A002 RI-123 Common Wrapper for Searching (Jude Saldanha, Dec-12-2023)
        private IRange FindNextMatch(IRange cells, string columnStr, int startRow, string term, LookAt lookAt)
        {
            int column = CellsHelper.ColumnNameToIndex(columnStr);
            return cells[1, column].EntireColumn.Find(
                    term, cells[startRow, column], FindLookIn.Values, lookAt,
                    SearchOrder.ByRows, SearchDirection.Next, false);
        }

        /// RI-A002 RI-123 Common Wrapper for Searching (Jude Saldanha, Dec-12-2023)
        private IRange FindNextPartialMatch(IRange cells, string columnStr, int startRow, string term)
        {
            return FindNextMatch(cells, columnStr, startRow, term, LookAt.Part);
        }

        /// RI-A002 RI-123 Common Wrapper for Searching (Jude Saldanha, Dec-12-2023)
        private IRange FindNextExactMatch(IRange cells, string columnStr, int startRow, string term)
        {
            return FindNextMatch(cells, columnStr, startRow, term, LookAt.Whole);
        }

        /// RI-A002 RI-123 Create the border for the total parts of the statement (Jude Saldanha, Dec-07-2023)
        private void CreateBorderTotal(IWorksheet worksheet)
        {
            // Find next row with total in Column C
            // Start from row 7
            int searchStartRow = 7;

            // Keep searching until we get to the end of the Used Range
            while (searchStartRow <= worksheet.UsedRange.Range.RowCount)
            {
                // Search in Column for next match
                IRange result = FindNextExactMatch(worksheet.Cells, "C", searchStartRow, "Total");

                // Terminate loop if nothing found or we looped back round to the beginning
                if (result == null || result.Row < searchStartRow)
                {
                    break;
                }

                // Next loop carry on from next row
                searchStartRow = result.Row;

                // If our Total is hidden, move on to next one
                if (result.EntireRow.Hidden)
                {
                    continue;
                }

                // End at Col G if H is Empty
                IRange colH = worksheet.Cells["H" + result.Row];
                string endCol = string.IsNullOrWhiteSpace(colH.Text) ? "H" : "G";

                int endRow = result.Row + 1;
                int startRow = result.Row;
                while (startRow >= 0 &&
                    string.IsNullOrWhiteSpace(worksheet.Cells["C" + startRow].Text))
                {
                    startRow--;
                }

                // Draw the border
                CreateBorderBlackMedium(worksheet, "C" + startRow, endCol + endRow);
            }
        }

        /// RI-A002 RI-123 Move Part 12 to end of Part10 and sort it (Jude Saldanha, Dec-07-2023)
        private void InsertPart12(IWorksheet worksheet)
        {
            IRange cells = worksheet.Cells;
            int colA = CellsHelper.ColumnNameToIndex("A");
            //int colB = CellsHelper.ColumnNameToIndex("B");
            int colC = CellsHelper.ColumnNameToIndex("C");
            int colH = CellsHelper.ColumnNameToIndex("H");
            int minSearchStartRow = 7;

            // Search in Column B for first match of part 12 
            IRange part12Start = FindNextExactMatch(worksheet.Cells, "B", minSearchStartRow, "Part 12");

            // Return if nothing found
            if (part12Start == null)
            {
                return;
            }

            // Search for first match of part 13 after part 12
            IRange part13Start = FindNextExactMatch(worksheet.Cells, "B", minSearchStartRow, "Part 13");

            // Return if nothing found
            if (part13Start == null)
            {
                return;
            }

            // Part 12 length - subtract 2 to exclude part12 headers and line between part 12 and part 13
            int part12len = part13Start.Row - part12Start.Row - 2;

            // Search for first match of part 11
            IRange part11Start = FindNextExactMatch(worksheet.Cells, "B", minSearchStartRow, "Part 11");

            // Part 10 end is 2 rows above part 11
            int part10EndRow = part11Start.Row - 2;

            // Insert blank rows under part 10 to make space for part 12
            int insertStart = part10EndRow + 2;
            int insertEnd = insertStart + part12len;
            for (int i = insertStart; i <= insertEnd; i++)
            {
                cells[i, 0].EntireRow.Insert();
            }

            // Copy rows over - add 2 to miss out header and add part12len as we've inserted that many rows above
            int copyStartRow = part12Start.Row + 2 + part12len;
            int copyEndRow = copyStartRow + part12len;
            cells[copyStartRow, colC, copyEndRow, colH].Copy(cells[insertStart - 1, colC, insertEnd - 1, colH]);
            part10EndRow = insertEnd;

            // Delete part 12, note again it has shifted down
            int deleteStart = part12Start.Row + 1 + part12len;
            int deleteEnd = deleteStart + part12len;
            for (int i = deleteStart; i <= deleteEnd; i++)
            {
                // Note the cells shift up on delete so always use start
                cells[deleteStart, 0].EntireRow.Delete();
            }

            // Get Part 10 start
            // Search for first match of part 10
            IRange part10Start = FindNextExactMatch(worksheet.Cells, "B", minSearchStartRow, "Part 10");

            // Get the Part 10 range to sort
            SpreadsheetGear.IRange range = worksheet.Cells[part10Start.Row + 1, colA, part10EndRow, colH];

            // Set up the first sort key with a key index of column c, ascending, normal
            SpreadsheetGear.SortKey sortKey1 = new SpreadsheetGear.SortKey(
            colC, SpreadsheetGear.SortOrder.Ascending, SpreadsheetGear.SortDataOption.Normal);

            // Sort the range by rows, ignoring case, passing the sort key.
            // NOTE: Any number of sort keys may be passed to the Sort method.
            range.Sort(SpreadsheetGear.SortOrientation.Rows, false, sortKey1);
        }

        /// RI-A002 RI-123 Page Break after Part 9 using SpreadsheetGear (Jude Saldanha, Dec-19-2023)
        private void PageBreakAfterPart9(IWorksheet worksheet)
        {
            IRange cells = worksheet.Cells;
            int minSearchStartRow = 7;

            // Search in Column B for first match of part 10
            IRange part10Start = FindNextExactMatch(worksheet.Cells, "B", minSearchStartRow, "Part 10");

            // Return if nothing found
            if (part10Start == null)
            {
                return;
            }

            worksheet.Cells[part10Start.Row - 1, 0].PageBreak = SpreadsheetGear.PageBreak.Manual;
        }

        /// RI-A002 RI-123 Page Break after Part 9 using Aspose (Jude Saldanha, Dec-19-2023)
        private void PageBreakAfterPart9(Worksheet worksheet)
        {
            // Find Part 10 and add a page break before
            FindOptions findOptions = new FindOptions();
            CellArea ca = new CellArea();
            ca.StartRow = 0;
            ca.StartColumn = 0;
            ca.EndRow = worksheet.Cells.MaxDataRow;
            ca.EndColumn = worksheet.Cells.MaxDataColumn;
            findOptions.SetRange(ca);
            findOptions.CaseSensitive = false;
            Cell cell = worksheet.Cells.Find("Part 10", null, findOptions);
            if (cell != null)
            {
                worksheet.HorizontalPageBreaks.Add(cell.Row - 1, 0);
            }
        }

        /// RI-A002 RI-123 Create borders around parts 10 and 11 (Jude Saldanha, Dec-19-2023)
        private void CreateBorderAroundPart(IWorksheet worksheet, string part)
        {
            IRange cells = worksheet.Cells;
            int minSearchStartRow = 7;

            IRange partStart = FindNextExactMatch(worksheet.Cells, "B", minSearchStartRow, part);
            if (partStart == null)
            {
                return;
            }

            int partEnd = partStart.Row + 2;
            while (!string.IsNullOrEmpty(cells["C" + partEnd].Text))
            {
                partEnd++;
            }


            CreateBorderBlackMedium(worksheet, "C" + (partStart.Row + 2), "H" + (partEnd - 1));
        }

        /// RI-A002 RI-123 Get the logo (Jude Saldanha, Dec-18-2023)
        private void GetLogo(IWorksheet worksheet)
        {
            // Get the UDF value from AA1, if not set default to alliantLogo
            string udfValue = worksheet.Cells["AA1"].Text;
            
            switch(udfValue)
            {
                case "alliantLogo":
                default:
                    worksheet.Shapes["Picture 1"].Visible = true;
                    worksheet.Shapes["Picture 2"].Visible = false;
                    worksheet.Shapes["Picture 3"].Visible = false;
                    break;

                case "RSS":
                    worksheet.Shapes["Picture 1"].Visible = false;
                    worksheet.Shapes["Picture 2"].Visible = true;
                    worksheet.Shapes["Picture 3"].Visible = false;
                    break;

                case "alliantText":
                    worksheet.Shapes["Picture 1"].Visible = false;
                    worksheet.Shapes["Picture 2"].Visible = false;
                    worksheet.Shapes["Picture 3"].Visible = true;
                    break;
            } 
        }


        /// RI-A002 RI-123 Get the logo, using a dictionary rather than case statement (Jude Saldanha, Dec-18-2023) <summary>
        /// </summary>
        /// <param name="worksheet"></param>
        private void GetLogoNew(IWorksheet worksheet)
        {
            // Create a mapping for logo UDF value to picture
            Dictionary<string, string> logoMap = new Dictionary<string, string>
            {
                {"alliantLogo", "Picture 1"},
                {"RSS", "Picture 2"},
                {"alliantText", "Picture 3"},
            };

            // Default all the logos to not visible
            foreach (string picture in logoMap.Values)
            {
                worksheet.Shapes[picture].Visible = false;
            }

            // Get the UDF value from AA1, if not set default to alliantLogo
            string udfValue = worksheet.Cells["AA1"].Text;
            string logoName = (logoMap.ContainsKey(udfValue)) ? logoMap[udfValue] : "Picture 1";


            // Set logo in UDF to visible
            if (worksheet.Shapes[logoName] != null)
            {
                worksheet.Shapes[logoName].Visible = true;
            }
        }

        private void SetStatementField(Workbook asposeWorkbook, string fieldName, string toValue)
        {
            if (toValue == null)
                toValue = "";

            var containsKey = asposeWorkbook.CustomDocumentProperties.Contains(fieldName);

            if (containsKey)
            {
                asposeWorkbook.CustomDocumentProperties[fieldName].Value = toValue;
                return;
            }

            asposeWorkbook.CustomDocumentProperties.Add(fieldName, toValue);
        }

        /// <summary>
        /// Excel has a function Range(Range1, Range2) that returns a Range from Range1 TO Range2. NOT a union.
        /// SpreadsheetGear does not appear to have this functionality so this is to be used as a replacement for it.
        /// </summary>
        /// <param name="range1">From Range</param>
        /// <param name="range2">To Range</param>
        private IRange RangeToRange(IRange range1, IRange range2)
        {
            var lastCell = range2[range2.RowCount - 1, range2.ColumnCount - 1];
            return range1.Worksheet.Range[range1.Row, range1.Column, lastCell.Row, lastCell.Column];
        }
    }
}
