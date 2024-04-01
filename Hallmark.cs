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
