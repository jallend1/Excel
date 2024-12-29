Sub ConsolidateData()
    Dim ws As Worksheet
    Dim wsConsolidate As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim nextRow As Long
    Dim sheetExists As Boolean
    
    ' Check if "ConsolidatedData" sheet exists
    sheetExists = False
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "ConsolidatedData" Then
            sheetExists = True
            Set wsConsolidate = ws
            Exit For
        End If
    Next ws
    
    ' If "ConsolidatedData" sheet does not exist, create it
    If Not sheetExists Then
        Set wsConsolidate = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsConsolidate.Name = "ConsolidatedData"
    End If
    
    nextRow = 2 ' Assuming row 1 has headers
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "ConsolidatedData" Then
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Copy wsConsolidate.Cells(nextRow, 1)
            nextRow = nextRow + lastRow - 1
        End If
    Next ws
End Sub