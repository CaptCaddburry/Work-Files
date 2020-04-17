'This is for adding a timestamp to a cooresponding cell when data is added to a given cell

Private Sub Worksheet_Change(ByVal Target As Range)
Dim xCellColumn As Integer
Dim xTimeColumn As Integer
Dim xRow, xCol As Integer
Dim xDPRg, xRg As Range
xCellColumn = 3
xTimeColumn = 5
xRow = Target.Row
xCol = Target.Column
If Target.Text <> "" Then
    If xCol = xCellColumn Then
       Cells(xRow, xTimeColumn) = Now()
    Else
        On Error Resume Next
        Set xDPRg = Target.Dependents
        For Each xRg In xDPRg
            If xRg.Column = xCellColumn Then
                Cells(xRg.Row, xTimeColumn) = Now()
            End If
        Next
    End If
End If
End Sub

'This is for saving a copy of the current file when the macro is run

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ThisWorkbook.Sheets.Copy
    DateStamp = Format(Date, "MM-DD-YYYY")
    With ActiveWorkbook
        .SaveAs ThisWorkbook.Path & "/" & DateStamp & " " & "filename.xlsx"
        .Close
    End With
End Sub
