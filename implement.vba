Sub Unfiltered()
'
' Unfiltered Macro
' To remove the filter.
'

'
    Range("A6").Select
    Selection.AutoFilter
End Sub
Sub Clear()
'
' Clear Macro
' To clear everything
'

'
    Range("A7").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("A7").Select
End Sub
Sub Open_Points()
'
' Open_Points Macro
' To filter out open points
'

'
    Range("I6").Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    ActiveSheet.Range("$A$6:$L$90").AutoFilter Field:=9, Criteria1:="Open"
End Sub

Sub RefreshMOMDashboard()

    Worksheets("MOMSummary").Activate
        'To unfilter the MOM summary table
        Range("A6").Select
        Selection.AutoFilter
    
        'To Clear the MOM Summary Table
        Range("A7").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.ClearContents
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        Selection.Borders(xlEdgeLeft).LineStyle = xlNone
        Selection.Borders(xlEdgeTop).LineStyle = xlNone
        Selection.Borders(xlEdgeBottom).LineStyle = xlNone
        Selection.Borders(xlEdgeRight).LineStyle = xlNone
        Selection.Borders(xlInsideVertical).LineStyle = xlNone
        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
        Range("A7").Select

    ' To count sheets in excel file
    totalsheets = Worksheets.Count

    For i = 5 To totalsheets

        'Checking last filled row on each sheet
        lastrow = Worksheets(i).Cells(Rows.Count, 1).End(xlUp).Row
    
                      For j = 10 To lastrow
                      Worksheets(i).Activate
                      Worksheets(i).Rows(j).Select
                      Selection.Copy
    
                      Worksheets("MOMSummary").Activate
        
                      lastrow = Worksheets("MOMSummary").Cells(Rows.Count, 1).End(xlUp).Row
        
                      Worksheets("MOMSummary").Cells(lastrow + 1, 1).Select
                      ActiveSheet.Paste
                      Next
    
    Next
    Worksheets("MOMDashboard").Activate
    Worksheets("MOMDashboard").Cells(1, 1).Select
    ThisWorkbook.RefreshAll
End Sub

Sub RefreshAttendanceDashboard()
' To add a new column with the name of the sheets

Dim colname As String
Dim dte As Date
Dim cell_value As Variant
Dim last_col As Long
Dim last_row As Long
Dim foundRng As Range
Dim tm_name As String
Dim rng As Range
Dim condition1 As FormatCondition, condition2 As FormatCondition, condition3 As FormatCondition
Dim iCells As Range


    'Getting name of the sheet
    colname = Sheets(Sheets.Count).Name
    'Converting name of heet to Date
    dte = CDate(colname)

    Worksheets("AttendanceSummary").Activate
        last_col = Cells(2, Columns.Count).End(xlToLeft).Column
        cell_value = Cells(2, Columns.Count).End(xlToLeft).Value
    
        'Checking if Date column exists
        'If not create a new table
        If dte <> CDate(cell_value) Then
            Cells(2, last_col + 1) = dte
            Cells(2, last_col + 1).Font.Bold = True
            'To align center
            Cells(2, last_col + 1).HorizontalAlignment = xlCenter
            
            last_row = Cells(Rows.Count, 2).End(xlUp).Row
        
            'If date is smaller than today's date
            'Populating data into the column
            If CDate(Cells(1, 5)) >= CDate(Cells(2, last_col + 1)) Then
                For j = 3 To last_row
                    tm_name = Cells(j, 2).Value
                    Set foundRng = Worksheets(colname).Range("H1:M8").Find(tm_name)
                    Cells(j, last_col + 1) = Worksheets(colname).Cells(foundRng.Row, foundRng.Column + 1)
                    Cells(j, last_col + 1).HorizontalAlignment = xlCenter
                    Cells(j, 3) = WorksheetFunction.CountIf(Worksheets("AttendanceSummary").Range(Cells(j, 6), Cells(j, last_col + 1)), "Yes")
                    Cells(j, 4) = WorksheetFunction.CountIf(Worksheets("AttendanceSummary").Range(Cells(j, 6), Cells(j, last_col + 1)), "No")
                    Cells(j, 5) = WorksheetFunction.CountIf(Worksheets("AttendanceSummary").Range(Cells(j, 6), Cells(j, last_col + 1)), "Unable to Attend")
                    
                Next
            End If
        End If
    
        'Formatting
        Set rng = Range("attendanceSummary")
        rng.FormatConditions.Delete
    
        'Defining and setting the criteria for each conditional format
        Set condition1 = rng.FormatConditions.Add(Type:=xlTextString, TextOperator:=xlContains, String:="Unable to Attend")
        Set condition2 = rng.FormatConditions.Add(Type:=xlTextString, TextOperator:=xlContains, String:="No")
        Set condition3 = rng.FormatConditions.Add(Type:=xlTextString, TextOperator:=xlContains, String:="Yes")

        'Defining and setting the format to be applied for each condition
        With condition1
            .Font.Color = RGB(156, 87, 0)
            .Interior.Color = RGB(255, 235, 156)
        End With

        With condition2
            .Font.Color = RGB(156, 0, 6)
            .Interior.Color = RGB(255, 199, 206)
        End With
   
        With condition3
            .Font.Color = RGB(0, 97, 0)
            .Interior.Color = RGB(198, 239, 206)
        End With
    
        'For Borders
        For Each iCells In rng
            iCells.BorderAround _
                LineStyle:=xlContinuous, _
                Weight:=xlThin
        Next iCells
Worksheets("AttendanceDashboard").Activate
Worksheets("AttendanceDashboard").Cells(1, 1).Select
ThisWorkbook.RefreshAll

End Sub

Private Sub CommandButton1_Click()

' To add a new column with the name of the sheets

Dim colname As String
Dim dte As Date
Dim cell_value As Variant
Dim last_col As Long
Dim last_row As Long
Dim foundRng As Range
Dim tm_name As String
Dim rng As Range
Dim condition1 As FormatCondition, condition2 As FormatCondition, condition3 As FormatCondition
Dim iCells As Range


    'Getting name of the sheet
    colname = Sheets(Sheets.Count).Name
    'Converting name of heet to Date
    dte = CDate(colname)

    Worksheets("AttendanceSummary").Activate
        last_col = Cells(2, Columns.Count).End(xlToLeft).Column
        cell_value = Cells(2, Columns.Count).End(xlToLeft).Value
    
        'Checking if Date column exists
        'If not create a new table
        If dte <> CDate(cell_value) Then
            Cells(2, last_col + 1) = dte
            Cells(2, last_col + 1).Font.Bold = True
            'To align center
            Cells(2, last_col + 1).HorizontalAlignment = xlCenter
            last_row = Cells(Rows.Count, 2).End(xlUp).Row
        
            'If date is smaller than today's date
            'Populating data into the column
            If CDate(Cells(1, 5)) >= CDate(Cells(2, last_col + 1)) Then
                For j = 3 To last_row
                    tm_name = Cells(j, 2).Value
                    Set foundRng = Worksheets(colname).Range("H1:M8").Find(tm_name)
                    Cells(j, last_col + 1) = Worksheets(colname).Cells(foundRng.Row, foundRng.Column + 1)
                    Cells(j, last_col + 1).HorizontalAlignment = xlCenter
                    Cells(j, 3) = WorksheetFunction.CountIf(Worksheets("AttendanceSummary").Range(Cells(j, 6), Cells(j, last_col + 1)), "Yes")
                    Cells(j, 4) = WorksheetFunction.CountIf(Worksheets("AttendanceSummary").Range(Cells(j, 6), Cells(j, last_col + 1)), "No")
                    Cells(j, 5) = WorksheetFunction.CountIf(Worksheets("AttendanceSummary").Range(Cells(j, 6), Cells(j, last_col + 1)), "Unable to Attend")
                    
                Next
            End If
        End If
    
        'Formatting
        Set rng = Range("attendanceSummary")
        rng.FormatConditions.Delete
    
        'Defining and setting the criteria for each conditional format
        Set condition1 = rng.FormatConditions.Add(Type:=xlTextString, TextOperator:=xlContains, String:="Unable to Attend")
        Set condition2 = rng.FormatConditions.Add(Type:=xlTextString, TextOperator:=xlContains, String:="No")
        Set condition3 = rng.FormatConditions.Add(Type:=xlTextString, TextOperator:=xlContains, String:="Yes")

        'Defining and setting the format to be applied for each condition
        With condition1
            .Font.Color = RGB(156, 87, 0)
            .Interior.Color = RGB(255, 235, 156)
        End With

        With condition2
            .Font.Color = RGB(156, 0, 6)
            .Interior.Color = RGB(255, 199, 206)
        End With
   
        With condition3
            .Font.Color = RGB(0, 97, 0)
            .Interior.Color = RGB(198, 239, 206)
        End With
    
        'For Borders
        For Each iCells In rng
            iCells.BorderAround _
                LineStyle:=xlContinuous, _
                Weight:=xlThin
        Next iCells
ThisWorkbook.RefreshAll
End Sub

Private Sub CommandButton1_Click()
' To count sheets in excel file
totalsheets = Worksheets.Count

For i = 5 To totalsheets

        'Checking last filled row on each sheet
        lastrow = Worksheets(i).Cells(Rows.Count, 1).End(xlUp).Row
    
                      For j = 10 To lastrow
                      Worksheets(i).Activate
                      Worksheets(i).Rows(j).Select
                      Selection.Copy
                      Worksheets("MOMSummary").Activate
        
                      Worksheets("MOMSummary").Activate
        
                      lastrow = Worksheets("MOMSummary").Cells(Rows.Count, 1).End(xlUp).Row
        
                      Worksheets("MOMSummary").Cells(lastrow + 1, 1).Select
                      ActiveSheet.Paste
                      Next
    
Next
Worksheets("MOMSummary").Cells(1, 1).Select
ThisWorkbook.RefreshAll
End Sub






