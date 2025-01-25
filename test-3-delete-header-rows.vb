Sub DeleteFilteredVisibleRows()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Worksheets("special") ' Adjust the worksheet name
    
    ' Check if filters are active
    If Not ws.AutoFilterMode Then
        MsgBox "No filters applied!", vbExclamation
        Exit Sub
    End If
    
    ' Get the visible range excluding the header row
    Set rng = ws.AutoFilter.Range.SpecialCells(xlCellTypeVisible)
    
    ' Loop through visible rows and delete those with "PARENT_NM"
    For Each cell In rng.Columns(1).Cells
        If cell.Row > 1 Then ' Skip the header row
            If cell.Value = "PARENT_NM" Then
                ws.Rows(cell.Row).Delete
            End If
        End If
    Next cell
End Sub
