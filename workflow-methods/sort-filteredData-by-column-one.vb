Sub SortDataByColumn1()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim i As Long
    Dim blankCount As Long
    
    ' Set the active worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("FilteredData")
    On Error GoTo 0

    If Not ws Is Nothing Then
        MsgBox "Tried to sort `FilteredData` but the sheet does not exist", vbExclamation
        Set ws = Nothing
        Exit Sub
    End If
    
    ' Set the starting row (skip the header row)
    startRow = 2
    
    ' Find the last row of data in column 1
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Initialize variables for blank cell tracking
    endRow = startRow
    blankCount = 0
    
    ' Loop through column 1 to find the first two consecutive blank cells
    For i = startRow To lastRow
        If ws.Cells(i, 1).Value = "" Then
            blankCount = blankCount + 1
            If blankCount = 2 Then
                endRow = i - 1 ' End the range at the row before the second blank cell
                Exit For
            End If
        Else
            blankCount = 0 ' Reset blank count if non-blank value is found
        End If
    Next i
    
    ' If no blanks were found, set endRow to the last row
    If blankCount < 2 Then
        endRow = lastRow
    End If
    
    ' Sort the data in column 1 from row 2 to the row before the first two consecutive blanks
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, 1)), Order:=xlAscending
        .SetRange ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, ws.Columns.Count).End(xlToLeft))
        .Header = xlNo ' Skip the header row
        .Apply
    End With
End Sub
