class SurroundingClass
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
        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo ErrorRoutine' at character 532
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 41
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorGoToStatement(OnErrorGoToStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorGoToStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
    On Error GoTo ErrorRoutine

 */
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

        TitleOffset = 0;
        TitleNewColK = 11;     // Col K
        TitleNewColL = 12;     // Col L
        TitleCurrentColM = 13; // Col M
        TitleNewColN = 14;     // Col N
        TitleNewColO = 15;     // Col O

        SearchColumn = 3; // Col C
        SearchStartRow = 1;
        SearchTerm = "TOTALS";


        {
            var withBlock = ThisWorkbook.Worksheets("Sheet1");
            ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 1796
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 41
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
                   
        'Find row in column C containing the "TOTALS"
        Set Result = .Columns(SearchColumn).Find(what:=SearchTerm, After:=.Cells(SearchStartRow, SearchColumn), LookIn:=xlValues, LookAt:= _
            xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)

 */
            if (!Result == null)
            {
                // Found the row with "TOTALS"
                TotalsRow = Result.Row;
                // Hide Units Returned if none displayed.
                if ((withBlock.Cells(TotalsRow, Col_UnitsReturned) == 0))
                {
                    withBlock.Columns(Col_UnitsReturned).EntireColumn.Hidden = true;
                    TitleOffset = TitleOffset + 1;
                }

                // Hide Per Item Rate if none displayed.
                if ((withBlock.Cells(TotalsRow, Col_PerItemRate) == 0))
                {
                    withBlock.Columns(Col_PerItemRate).EntireColumn.Hidden = true;
                    TitleOffset = TitleOffset - 1;
                }
                else
                    // Grandtotal is not required for Per Item Rate column.
                    withBlock.Cells(TotalsRow, Col_PerItemRate) = "";

                // HALS-49 Contributor Share / Contributor Royalty Earned are unconditionally hidden
                TitleOffset = TitleOffset - 2;
                // Grandtotal is not required for Contributor Share column.
                withBlock.Cells(TotalsRow, Col_ContributorShare) = "";

                // Hide Reserves Taken if none displayed.
                if ((withBlock.Cells(TotalsRow, Col_ReservesTaken) == 0))
                {
                    withBlock.Columns(Col_ReservesTaken).EntireColumn.Hidden = true;
                    TitleOffset = TitleOffset - 1;
                }

                // Hide Reserves Liquidated if none displayed.
                if ((withBlock.Cells(TotalsRow, Col_ReservesLiquidated) == 0))
                {
                    withBlock.Columns(Col_ReservesLiquidated).EntireColumn.Hidden = true;
                    TitleOffset = TitleOffset - 1;
                }
            }

            // Move the main statement title (currently in col M) if needed.
            if (TitleOffset != 0)
            {
                // Copy the title
                withBlock.Range(withBlock.Cells(1, TitleCurrentColM), withBlock.Cells(5, TitleCurrentColM)).Copy();

                // Paste title
                if (TitleOffset < -2)
                {
                    // find unhidden column
                    while (withBlock.Columns(TitleNewColK).Hidden == true)
                        TitleNewColK = TitleNewColK - 1;
                    withBlock.Range(withBlock.Cells(1, TitleNewColK), withBlock.Cells(5, TitleNewColK)).PasteSpecial(Paste: xlValues);
                    withBlock.Range(withBlock.Cells(1, TitleNewColK), withBlock.Cells(5, TitleNewColK)).PasteSpecial(Paste: xlPasteFormats);
                }
                else if (TitleOffset < 0)
                {
                    // find unhidden column
                    while (withBlock.Columns(TitleNewColL).Hidden == true)
                        TitleNewColL = TitleNewColL - 1;
                    withBlock.Range(withBlock.Cells(1, TitleNewColL), withBlock.Cells(5, TitleNewColL)).PasteSpecial(Paste: xlValues);
                    withBlock.Range(withBlock.Cells(1, TitleNewColL), withBlock.Cells(5, TitleNewColL)).PasteSpecial(Paste: xlPasteFormats);
                }
                else if (TitleOffset > 0)
                {
                    // find unhidden column
                    while (withBlock.Columns(TitleNewColN).Hidden == true)
                        TitleNewColN = TitleNewColN - 1;
                    withBlock.Range(withBlock.Cells(1, TitleNewColN), withBlock.Cells(5, TitleNewColN)).PasteSpecial(Paste: xlValues);
                    withBlock.Range(withBlock.Cells(1, TitleNewColN), withBlock.Cells(5, TitleNewColN)).PasteSpecial(Paste: xlPasteFormats);
                }
                else if (TitleOffset > 2)
                {
                    // find unhidden column
                    while (withBlock.Columns(TitleNewColO).Hidden == true)
                        TitleNewColO = TitleNewColO - 1;
                    withBlock.Range(withBlock.Cells(1, TitleNewColO), withBlock.Cells(5, TitleNewColO)).PasteSpecial(Paste: xlValues);
                    withBlock.Range(withBlock.Cells(1, TitleNewColO), withBlock.Cells(5, TitleNewColO)).PasteSpecial(Paste: xlPasteFormats);
                }

                // Clear original title.
                withBlock.Range(withBlock.Cells(1, TitleCurrentColM), withBlock.Cells(5, TitleCurrentColM)).Clear();
            }

            // Move Recoupment Group from column B into column C.
            TotalsRow = withBlock.Range(Col_C + withBlock.Rows.Count).End(xlUp).Row;
            // Going through all the rows
            for (Counter = 1; Counter <= TotalsRow; Counter++)
            {
                // Search for RG@@ text in column B.
                if ((InStr(withBlock.Cells(Counter, Col_B), "RG@@") > 0))
                {
                    // Recoupment Group found, copy to column C
                    withBlock.Range(withBlock.Cells(Counter, Col_B), withBlock.Cells(Counter, Col_B)).Copy();
                    withBlock.Range(withBlock.Cells(Counter, Col_C), withBlock.Cells(Counter, Col_C)).PasteSpecial(Paste: xlValues);
                    withBlock.Range(withBlock.Cells(Counter, Col_C), withBlock.Cells(Counter, Col_C)).PasteSpecial(Paste: xlPasteFormats);
                    // Strip out RG@@
                    withBlock.Cells(Counter, Col_C) = Mid(withBlock.Cells(Counter, Col_C), 5);
                    // Blank column B
                    withBlock.Cells(Counter, Col_B) = "";
                }
            }
        }

    ErrorRoutine:
        ;
        ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 6763
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 41
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 

    Set Result = Nothing

 */
    }
    // Hide "Net Roylaties Earned" column if all values are same for "Royalties Earned" column
    // HAL-A002 HAL-92\HAL-202 Multiple fields change (Peter Jun, Oct-01-2018)
    private void HideNetRoyaltiesEarned()
    {
        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo ErrorRoutine' at character 6998
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 41
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorGoToStatement(OnErrorGoToStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorGoToStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
On Error GoTo ErrorRoutine

 */
        long FirstRow;
        long LastRow;
        long Counter;
        int TargetCol;
        int HideCol;
        int SameFlag;

        {
            var withBlock = ThisWorkbook.Worksheets("Sheet1");
            SameFlag = 0;
            FirstRow = 9;
            TargetCol = 18;  // column R
            HideCol = 23;    // column W
            LastRow = withBlock.Cells(Rows.Count, TargetCol).End(xlUp).Row;

            for (Counter = FirstRow; Counter <= LastRow; Counter++)
            {
                if (withBlock.Cells(Counter, TargetCol) != withBlock.Cells(Counter, HideCol))
                    SameFlag = 1;
            }

            if (SameFlag == 0)
                withBlock.Columns(HideCol).Hidden = true;
        }

    ErrorRoutine:
        ;
    }
    // Hide logo depend on Company UDF value
    // HAL-A003 HAL-323\HAL-332 DaySpring Update (Peter Jun, Jan-07-2020)
    private void HidingLogo()
    {
        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo ErrRoutine' at character 7805
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 41
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorGoToStatement(OnErrorGoToStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorGoToStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
    On Error GoTo ErrRoutine

 */
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

    ErrRoutine:
        ;
    }
}
