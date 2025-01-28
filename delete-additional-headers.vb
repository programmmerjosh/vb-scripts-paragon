Sub DeleteAdditionalHeaders()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowNum As Long
    Dim firstHeaderFound As Boolean
    Dim isHeaderRow As Boolean
    Dim col As Range
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Worksheets("special") ' Adjust as needed
    
    ' Turn off filters to ensure proper deletion
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
    
    ' Get the last used row in the sheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Initialize the flag to track the first header row
    firstHeaderFound = False
    
    ' Loop through each row
    For rowNum = 1 To lastRow
        ' Assume this row is not a header
        isHeaderRow = False
        
        ' Check if this row is a header by looking for filters/dropdowns in columns
        For Each col In ws.Rows(rowNum).Columns
            If Not IsEmpty(col.Value) And col.Validation.Type = xlValidateList Then
                isHeaderRow = True
                Exit For
            End If
        Next col
        
        ' Handle the row based on whether it's a header
        If isHeaderRow Then
            If Not firstHeaderFound Then
                ' Keep the first header row
                firstHeaderFound = True
            Else
                ' Delete all subsequent header rows
                ws.Rows(rowNum).Delete
                rowNum = rowNum - 1 ' Adjust row index after deletion
                lastRow = lastRow - 1 ' Update the last row
            End If
        End If
    Next rowNum
    
    MsgBox "All additional headers removed, keeping only the first header row.", vbInformation
End Sub
