Sub DeleteFilteredRows()
    Dim ws As Worksheet
    Dim rng As Range
    Dim searchRange As Range
    Dim foundCell As Range
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Worksheets("special") ' Adjust the worksheet name
    
        ' Turn off filters to ensure proper deletion
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    Else
        MsgBox "No filters applied!", vbExclamation
        Exit Sub
    End If
    
    ' Define the range to search in (exclude the header row)
    Set searchRange = ws.Range("A2:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)
    
    ' Find and delete rows containing "PARENT_NM"
    Set foundCell = searchRange.Find(What:="PARENT_NM", LookIn:=xlValues, LookAt:=xlWhole)
    
    Do While Not foundCell Is Nothing
        ws.Rows(foundCell.Row).Delete
        Set searchRange = ws.Range("A2:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row) ' Redefine search range
        Set foundCell = searchRange.Find(What:="PARENT_NM", LookIn:=xlValues, LookAt:=xlWhole)
    Loop
End Sub
