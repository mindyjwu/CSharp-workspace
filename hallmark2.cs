 // Install-Package Microsoft.VisualBasic

internal partial class SurroundingClass
{
    public void PostExecuteProcessing()
    {
        Application.ScreenUpdating = false;
        Application.Calculation = xlCalculationManual;

        FormatStatement();
        HideNetRoyaltiesEarned();
        HidingLogo();

        Application.Calculation = xlCalculationAutomatic;
        Application.ScreenUpdating = true;

    }
    // HAL-A002 HAL-92\HAL-202 Fixed disappearing header (Peter Jun, Oct-16-2018)
    // Initial Revision: HAL-A002 HAL-3 Initial Implementation (Kenny Lau, Oct-30-2017)
    private void FormatStatement()
    {
        try
        {
            Range Result;
            long TotalsRow;
            long Counter;
            string Col_A;
            string Col_B;
            string Col_C;
            string Col_UnitsReturned;
            string Col_PerItemRate;
            string Col_ContributorShare;
            string Col_ContributorRoyaltyEarned;
            string Col_ReservesTaken;
            string Col_ReservesLiquidated;
            long SearchStartRow;
            long SearchColumn;
            string SearchTerm;
            long TitleOffset;
            long TitleNewColK;
            long TitleNewColL;
            long TitleCurrentColM;
            long TitleNewColN;
            long TitleNewColO;

            Col_A = "A";
            Col_B = "B";
            Col_C = "C";
            Col_UnitsReturned = "H";
            Col_PerItemRate = "Q";
            Col_ContributorShare = "S";
            Col_ContributorRoyaltyEarned = "T";
            Col_ReservesTaken = "U";
            Col_ReservesLiquidated = "V";
            TotalsRow = -1;

            TitleOffset = 0L;
            TitleNewColK = 11L;     // Col K
            TitleNewColL = 12L;     // Col L
            TitleCurrentColM = 13L; // Col M
            TitleNewColN = 14L;     // Col N
            TitleNewColO = 15L;     // Col O

            SearchColumn = 3L; // Col C
            SearchStartRow = 1L;
            SearchTerm = "TOTALS";


            {
                var withBlock = ThisWorkbook.Worksheets("Sheet1");
                ;
#error Cannot convert EmptyStatementSyntax - see comment for details
                /* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 1884


                                Input:

                                        'Find row in column C containing the "TOTALS"
                                        Set Result = .Columns(SearchColumn).Find(what:=SearchTerm, After:=.Cells(SearchStartRow, SearchColumn), LookIn:=xlValues, LookAt:= _
                                            xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)

                                 */
                if (Result is not null)
                {
                    // Found the row with "TOTALS"
                    TotalsRow = Result.Row;
                    // Hide Units Returned if none displayed.
                    if (withBlock.Cells(TotalsRow, Col_UnitsReturned) == 0)
                    {
                        withBlock.Columns(Col_UnitsReturned).EntireColumn.Hidden = true;
                        TitleOffset = TitleOffset + 1L;
                    }

                    // Hide Per Item Rate if none displayed.
                    if (withBlock.Cells(TotalsRow, Col_PerItemRate) == 0)
                    {
                        withBlock.Columns(Col_PerItemRate).EntireColumn.Hidden = true;
                        TitleOffset = TitleOffset - 1L;
                    }
                    else
                    {
                        // Grandtotal is not required for Per Item Rate column.
                        withBlock.Cells(TotalsRow, Col_PerItemRate) = "";
                    }

                    // HALS-49 Contributor Share / Contributor Royalty Earned are unconditionally hidden
                    TitleOffset = TitleOffset - 2L;
                    // Grandtotal is not required for Contributor Share column.
                    withBlock.Cells(TotalsRow, Col_ContributorShare) = "";

                    // Hide Reserves Taken if none displayed.
                    if (withBlock.Cells(TotalsRow, Col_ReservesTaken) == 0)
                    {
                        withBlock.Columns(Col_ReservesTaken).EntireColumn.Hidden = true;
                        TitleOffset = TitleOffset - 1L;
                    }

                    // Hide Reserves Liquidated if none displayed.
                    if (withBlock.Cells(TotalsRow, Col_ReservesLiquidated) == 0)
                    {
                        withBlock.Columns(Col_ReservesLiquidated).EntireColumn.Hidden = true;
                        TitleOffset = TitleOffset - 1L;
                    }
                }

                // Move the main statement title (currently in col M) if needed.
                if (TitleOffset != 0L)
                {
                    // Copy the title
                    withBlock.Range(withBlock.Cells(1, TitleCurrentColM), withBlock.Cells(5, TitleCurrentColM)).Copy();

                    // Paste title
                    if (TitleOffset < -2)
                    {
                        // find unhidden column
                        while (withBlock.Columns(TitleNewColK).Hidden == true)
                            TitleNewColK = TitleNewColK - 1L;
                        withBlock.Range(withBlock.Cells(1, TitleNewColK), withBlock.Cells(5, TitleNewColK)).PasteSpecial(Paste: xlValues);
                        withBlock.Range(withBlock.Cells(1, TitleNewColK), withBlock.Cells(5, TitleNewColK)).PasteSpecial(Paste: xlPasteFormats);
                    }
                    else if (TitleOffset < 0L)
                    {
                        // find unhidden column
                        while (withBlock.Columns(TitleNewColL).Hidden == true)
                            TitleNewColL = TitleNewColL - 1L;
                        withBlock.Range(withBlock.Cells(1, TitleNewColL), withBlock.Cells(5, TitleNewColL)).PasteSpecial(Paste: xlValues);
                        withBlock.Range(withBlock.Cells(1, TitleNewColL), withBlock.Cells(5, TitleNewColL)).PasteSpecial(Paste: xlPasteFormats);
                    }
                    else if (TitleOffset > 0L)
                    {
                        // find unhidden column
                        while (withBlock.Columns(TitleNewColN).Hidden == true)
                            TitleNewColN = TitleNewColN - 1L;
                        withBlock.Range(withBlock.Cells(1, TitleNewColN), withBlock.Cells(5, TitleNewColN)).PasteSpecial(Paste: xlValues);
                        withBlock.Range(withBlock.Cells(1, TitleNewColN), withBlock.Cells(5, TitleNewColN)).PasteSpecial(Paste: xlPasteFormats);
                    }
                    else if (TitleOffset > 2L)
                    {
                        // find unhidden column
                        while (withBlock.Columns(TitleNewColO).Hidden == true)
                            TitleNewColO = TitleNewColO - 1L;
                        withBlock.Range(withBlock.Cells(1, TitleNewColO), withBlock.Cells(5, TitleNewColO)).PasteSpecial(Paste: xlValues);
                        withBlock.Range(withBlock.Cells(1, TitleNewColO), withBlock.Cells(5, TitleNewColO)).PasteSpecial(Paste: xlPasteFormats);
                    }

                    // Clear original title.
                    withBlock.Range(withBlock.Cells(1, TitleCurrentColM), withBlock.Cells(5, TitleCurrentColM)).Clear();
                }

                // Move Recoupment Group from column B into column C.
                TotalsRow = withBlock.Range(Col_C + withBlock.Rows.Count).End(xlUp).Row;
                // Going through all the rows
                var loopTo = TotalsRow;
                for (Counter = 1L; Counter <= loopTo; Counter++)
                {
                    // Search for RG@@ text in column B.
                    if (Strings.InStr(withBlock.Cells(Counter, Col_B), "RG@@") > 0)
                    {
                        // Recoupment Group found, copy to column C
                        withBlock.Range(withBlock.Cells(Counter, Col_B), withBlock.Cells(Counter, Col_B)).Copy();
                        withBlock.Range(withBlock.Cells(Counter, Col_C), withBlock.Cells(Counter, Col_C)).PasteSpecial(Paste: xlValues);
                        withBlock.Range(withBlock.Cells(Counter, Col_C), withBlock.Cells(Counter, Col_C)).PasteSpecial(Paste: xlPasteFormats);
                        // Strip out RG@@
                        withBlock.Cells(Counter, Col_C) = Strings.Mid(withBlock.Cells(Counter, Col_C), 5);
                        // Blank column B
                        withBlock.Cells(Counter, Col_B) = "";
                    }

                }

            }
        }
        catch
        {
        };
#error Cannot convert EmptyStatementSyntax - see comment for details
        /* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 6975


        Input:

            Set Result = Nothing

         */
    }
    // Hide "Net Roylaties Earned" column if all values are same for "Royalties Earned" column
    // HAL-A002 HAL-92\HAL-202 Multiple fields change (Peter Jun, Oct-01-2018)
    private void HideNetRoyaltiesEarned()
    {
        try
        {
            long FirstRow;
            long LastRow;
            long Counter;
            int TargetCol;
            int HideCol;
            int SameFlag;

            {
                var withBlock = ThisWorkbook.Worksheets("Sheet1");
                SameFlag = 0;
                FirstRow = 9L;
                TargetCol = 18;  // column R
                HideCol = 23;    // column W
                LastRow = withBlock.Cells(Rows.Count, TargetCol).End(xlUp).Row;

                var loopTo = LastRow;
                for (Counter = FirstRow; Counter <= loopTo; Counter++)
                {
                    if (withBlock.Cells(Counter, TargetCol) != withBlock.Cells(Counter, HideCol))
                        SameFlag = 1;
                }

                if (SameFlag == 0)
                    withBlock.Columns(HideCol).Hidden = true;
            }
        }
        catch
        {
        }

    }
    // Hide logo depend on Company UDF value
    // HAL-A003 HAL-323\HAL-332 DaySpring Update (Peter Jun, Jan-07-2020)
    private void HidingLogo()
    {
        try
        {
            int SearchCol;
            int SearchRow;

            SearchRow = 1;
            SearchCol = 53;  // column BA

            {
                var withBlock = ThisWorkbook.Worksheets("Sheet1");
                if (withBlock.Cells(SearchRow, SearchCol) == "Hallmark")
                {
                    withBlock.Pictures("Picture 1").Visible = true;
                    withBlock.Pictures("Picture 2").Visible = false;
                }
                else if (withBlock.Cells(SearchRow, SearchCol) == "DaySpring")
                {
                    withBlock.Pictures("Picture 1").Visible = false;
                    withBlock.Pictures("Picture 2").Visible = true;
                }
                else
                {
                    withBlock.Pictures("Picture 1").Visible = true;
                    withBlock.Pictures("Picture 2").Visible = false;
                }
            }
        }
        catch
        {
        }


    }

}
