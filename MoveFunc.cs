// 6. 'Move Main Statement from column M if  needed
      // Check if TitleOffset is not equal to 0
      //'Move the main statement title (currently in col M) if needed.
      //If TitleOffset<> 0 Then
      int colBNum = 1;
      int colCNum = 2;
      if (titleOffset != 0)
      {
		//'Copy the title
		//.Range(.Cells(1, TitleCurrentColM), .Cells(5, TitleCurrentColM)).Copy
		IRange sourceRange = worksheet.Cells[0, titleCurrentColM, 4, titleCurrentColM];
		sourceRange.FillDown();
		// Paste title based on TitleOffset
		//If TitleOffset< -2 Then
		//' find unhidden column
		//Do While .Columns(TitleNewColK).Hidden = True
		//	TitleNewColK = TitleNewColK - 1
		//Loop
		//.Range(.Cells(1, TitleNewColK), .Cells(5, TitleNewColK)).PasteSpecial Paste:= xlValues
		//.Range(.Cells(1, TitleNewColK), .Cells(5, TitleNewColK)).PasteSpecial Paste:= xlPasteFormats
	if (titleOffset < -2)
		{
			while (worksheet.Cells[totalsRow, colV].EntireColumn.Hidden == true)
			{
				titleNewColK--;
			}
			CopyAndPasteValuesAndFormats(worksheet, column);
		}
		else if (titleOffset < 0)
		{
			while (worksheet.Cells[titleNewColL].Columns.Hidden == true)
			{
				titleNewColL--;
			}
			CopyAndPasteValuesAndFormats(worksheet, column);
		}
	}
          else if (titleOffset < 0)
          else if (titleOffset > 0)
          else
          // Find unhidden column
          while (worksheet.Columns[newCol].Hidden)
          {
              newCol--;
          }
          // Paste the title
          worksheet.Cells["1:" + "5," + newCol].PasteSpecial(SpreadsheetGear.PasteType.Values);
          worksheet.Cells["1:" + "5," + newCol].PasteSpecial(SpreadsheetGear.PasteType.Formats);
          worksheet.Cells[counter, colCNum].Value = worksheet.Cells[counter, colBNum].Value;
          worksheet.Cells[counter, colCNum].Style = worksheet.Cells[counter, colBNum].Style;
          // Clear original title
          worksheet.Cells["1:" + "5," + titleCurrentColM].Clear();
          //7. 'Move Recoupment Group from column B into column C.
          //        TotalsRow = .Range(Col_C & .Rows.Count).End(xlUp).Row
          totalsRow = worksheet.Cells[worksheet.Cells.Rows.Count, colC].End(SpreadsheetGear.xlUp).Row;
          totalsRow = worksheet.Cells.ColumnCount(colC);
          //        'Going through all the rows
          //        For Counter = 1 To TotalsRow
          for (counter = 1; counter <= totalsRow; counter++)
          {
              //            'Search for RG@@ text in column B.
              //            If(InStr(.Cells(Counter, Col_B), "RG@@") > 0) Then
              if (worksheet.Cells[counter, colBNum].Value != null && worksheet.Cells[counter, colBNum].Value.ToString().Contains("RG@@");
              {
                  // Then
                  //                'Recoupment Group found, copy to column C
                  //                .Range(.Cells(Counter, Col_B), .Cells(Counter, Col_B)).Copy
                  worksheet.Cells[counter, colCNum].Copy(worksheet.Cells[counter, colBNum]);
                  //                .Range(.Cells(Counter, Col_C), .Cells(Counter, Col_C)).PasteSpecial Paste:= xlValues
                  worksheet.Cells[counter, colCNum].Value = worksheet.Cells[counter, colBNum].Value;
                  worksheet.Cells[counter, colCNum].Style = worksheet.Cells[counter, colBNum].Style;
                  //                .Range(.Cells(Counter, Col_C), .Cells(Counter, Col_C)).PasteSpecial Paste:= xlPasteFormats
                  //                'Strip out RG@@
                  //                .Cells(Counter, Col_C) = Mid(.Cells(Counter, Col_C), 5)
                  string newValue = CellsHelper.     ("RG@@", "");
                  worksheet.Cells[counter, colCNum].Value = newValue;
                  //                'Blank column B
                  //                .Cells(Counter, Col_B) = ""
                  worksheet.Cells[counter, colBNum].Value = "";
                  //            End If
              }
          }
      }
}
