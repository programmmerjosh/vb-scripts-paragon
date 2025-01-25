Sub DeleteTableRows()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Worksheets("special") ' Adjust the worksheet name
    
    ' Check if there is a table in column 1
    If ws.ListObjects.Count > 0 Then
        ' Assume the table you want is the first one
        Set tbl = ws.ListObjects(1)
        
        ' Loop through table rows (backward to avoid shifting issues)
        For i = tbl.ListRows.Count To 1 Step -1
            If tbl.ListRows(i).Range.Cells(1, 1).Value = "PARENT_NM" Then
                tbl.ListRows(i).Delete
            End If
        Next i
    Else
        MsgBox "No table found on the sheet!", vbExclamation
    End If
End Sub
