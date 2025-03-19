Attribute VB_Name = "mod_SCurves"
Sub GenerateSCurve()
    Dim totalBudget As Double
    Dim numOfMonths As Long
    Dim i As Long
    Dim j As Long
    Dim LCMOfMonthsAnd4 As Long
    Dim monthlyBudget As Double
    Dim quarter1 As Long, quarter2 As Long, quarter3 As Long, quarter4 As Long
    Dim quarter1Budget As Double, quarter2Budget As Double, quarter3Budget As Double, quarter4Budget As Double
    Dim selectedRange As Range
    Dim selectedCell As Range
    Dim LCMBudgetArray() As Double
    Dim finalBudgetArray() As Double
    
    
'************** CHANGE IF CHANGING SPREADSHEETS *************************
    
    ' Set the budget column & forecasted column
    Dim budgetColumn As Integer: budgetColumn = 42
    Dim forecastedColumn As Integer: forecastedColumn = 44
    
'************************************************************************
    
    ' Get the selected range
    Set selectedRange = Selection
    
    Application.ScreenUpdating = False
    
    ' Check if the selection is a range
    If TypeName(selectedRange) = "Range" Then
    
        ' Check if the column of the leftmost selected cell is after column 3 and below row 2
        If Not CheckSelection(selectedRange, 26, 4) Then
            MsgBox "Please select a range starting in the forecast section"
            Application.ScreenUpdating = True
            Exit Sub
        End If
    
        ' Check if the selected range has more than one row
        If selectedRange.Rows.Count > 1 Then
            MsgBox "Please select cells from a single row."
            Application.ScreenUpdating = True
            Exit Sub
        End If
    
        ' Get the row of the first cell in the selected range
        rowNum = selectedRange.Cells(1, 1).Row
        
        ' Get the total budget from column B for the same row as the selected range
        totalBudget = Cells(rowNum, budgetColumn).value
        
        ' Get the number of months from the number of columns in the selected range
        numOfMonths = selectedRange.Columns.Count
        
        ' Calculate the LCM of the number of months and 4
        LCMOfMonthsAnd4 = WorksheetFunction.Lcm(numOfMonths, 4)
        
        ' Calculate the number of months in each quarter
        quarter1 = WorksheetFunction.RoundUp(LCMOfMonthsAnd4 / 4, 0)
        quarter2 = WorksheetFunction.RoundUp(LCMOfMonthsAnd4 / 2, 0) - quarter1
        quarter3 = WorksheetFunction.RoundUp(3 * LCMOfMonthsAnd4 / 4, 0) - quarter2 - quarter1
        quarter4 = LCMOfMonthsAnd4 - quarter3 - quarter2 - quarter1
        
        ' Calculate the budget for each quarter
        quarter1Budget = totalBudget * 0.1
        quarter2Budget = totalBudget * 0.2
        quarter3Budget = totalBudget * 0.38
        quarter4Budget = totalBudget * 0.32
        
        ' Initialize LCMBudgetArray
        ReDim LCMBudgetArray(LCMOfMonthsAnd4 - 1)
        
        ' Calculate the monthly budget for each quarter and store in LCMBudgetArray
        For i = 1 To LCMOfMonthsAnd4
            If i <= quarter1 Then
                LCMBudgetArray(i - 1) = quarter1Budget / quarter1
            ElseIf i <= quarter1 + quarter2 Then
                LCMBudgetArray(i - 1) = quarter2Budget / quarter2
            ElseIf i <= quarter1 + quarter2 + quarter3 Then
                LCMBudgetArray(i - 1) = quarter3Budget / quarter3
            Else
                LCMBudgetArray(i - 1) = quarter4Budget / quarter4
            End If
        Next i
        
        ' Initialize finalBudgetArray
        ReDim finalBudgetArray(numOfMonths - 1)
        
        ' Calculate the final budget for each month
        For i = 1 To numOfMonths
            finalBudgetArray(i - 1) = 0
            For j = 1 To LCMOfMonthsAnd4 / numOfMonths
                finalBudgetArray(i - 1) = finalBudgetArray(i - 1) + LCMBudgetArray((i - 1) * LCMOfMonthsAnd4 / numOfMonths + j - 1)
            Next j
        Next i
        
        ' Output the final budget array to the spreadsheet
        i = 1
        For Each selectedCell In selectedRange
            selectedCell.value = finalBudgetArray(i - 1)
            selectedCell.Interior.Color = RGB(240, 255, 240) ' Light green color
            selectedCell.Borders.LineStyle = xlContinuous
            i = i + 1
        Next selectedCell
        
        ' Insert the SUM formula in column 17 of the selected row
        Cells(selectedRange.Cells(1, 1).Row, forecastedColumn).Formula = "=SUM(" & selectedRange.Address & ")"
    
    Else
        MsgBox "Please select a range of cells."
    End If
    
    Application.ScreenUpdating = True
    
End Sub

Sub ClearCells()
    Dim selectedRange As Range
    Dim selectedCell As Range
    
    ' Get the selected range
    Set selectedRange = Selection
    
    Application.ScreenUpdating = False
    
    ' Check if the selection is a range
    If TypeName(selectedRange) = "Range" Then
    
        ' Check if the column of the leftmost selected cell is after column 3 and below row 2
        If Not CheckSelection(selectedRange, 26, 4) Then
            MsgBox "Please select a range starting in the forecast section"
            Application.ScreenUpdating = True
            Exit Sub
        End If
    
        ' Clear the values and set the background color to white for each cell in the selected range
        For Each selectedCell In selectedRange
            selectedCell.ClearContents
            selectedCell.Interior.Color = RGB(255, 255, 255) ' White color
            selectedCell.Borders.LineStyle = xlContinuous
        Next selectedCell
    
    Else
        MsgBox "Please select a range of cells."
    End If
    
    Application.ScreenUpdating = True
    
End Sub


Sub LinearDistribution()
    Dim totalBudget As Double
    Dim constructionBudget As Double
    Dim designBudget As Double
    Dim numOfMonths As Long
    Dim i As Long
    Dim selectedRange As Range
    Dim selectedCell As Range
    

'************** CHANGE IF CHANGING SPREADSHEETS *************************
    
    ' Set the budget column & forecasted column
    Dim budgetColumn As Integer: budgetColumn = 41
    Dim forecastedColumn As Integer: forecastedColumn = 43
    
'************************************************************************
    
    ' Get the selected range
    Set selectedRange = Selection
    
    Application.ScreenUpdating = False
    
    ' Check if the selection is a range
    If TypeName(selectedRange) = "Range" Then
    
        ' Check if the column of the leftmost selected cell is after column 3 and below row 2
        If Not CheckSelection(selectedRange, 26, 4) Then
            MsgBox "Please select a range starting in the forecast section"
            Application.ScreenUpdating = True
            Exit Sub
        End If
    
        ' Check if the selected range has more than one row
        If selectedRange.Rows.Count > 1 Then
            MsgBox "Please select cells from a single row."
            Application.ScreenUpdating = True
            Exit Sub
        End If
    
        ' Get the row of the first cell in the selected range
        rowNum = selectedRange.Cells(1, 1).Row
        
        ' Get the total budget from column B for the same row as the selected range
        'constructionBudget = Cells(rowNum, 9).Value
        'totalBudget = Cells(rowNum, 5).Value
        designBudget = Cells(rowNum, budgetColumn).value
        ' Get the number of months from the number of cells in the selected range
        numOfMonths = selectedRange.Cells.Count
        
        ' Calculate the budget for each month
        monthlyBudget = designBudget / numOfMonths
        
        ' Output the final budget array to the spreadsheet
        i = 1
        For Each selectedCell In selectedRange
            selectedCell.value = monthlyBudget
            selectedCell.Interior.Color = RGB(240, 250, 255) ' Light blue color
            selectedCell.Borders.LineStyle = xlContinuous
            i = i + 1
        Next selectedCell
        
        ' Insert the SUM formula in column 17 of the selected row
        Cells(selectedRange.Cells(1, 1).Row, forecastedColumn).Formula = "=SUM(" & selectedRange.Address & ")"
    
    Else
        MsgBox "Please select a range of cells."
    End If
    
    Application.ScreenUpdating = True
    
End Sub

Function CheckSelection(selectedRange As Range, minColumn As Long, minRow As Long) As Boolean
    If selectedRange.Cells(1, 1).Column < minColumn Or selectedRange.Cells(1, 1).Row < minRow Then
        CheckSelection = False
    Else
        CheckSelection = True
    End If
End Function
