Sub MergeMySheets()
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim lastRow As Long, targetRow As Long
    Dim wb As Workbook
    Dim sheetNames As Variant
    Dim i As Integer
    
    ' Define the sheet names to copy from
    sheetNames = Array("s1", "s2", "s3") ' Update as needed
    
     ' Create a new worksheet for the filtered data
    Set wb = ThisWorkbook
    Set wsTarget = wb.Sheets.Add
    wsTarget.Name = "special" ' Change as needed

    ' Clear existing data in target sheet before merging
    wsTarget.Cells.Clear
    targetRow = 1 ' Start pasting from the first row

    ' Loop through each source sheet
    For i = LBound(sheetNames) To UBound(sheetNames)
        Set wsSource = wb.Sheets(sheetNames(i))
        
        ' Find last used row in source sheet
        lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
        
        ' Copy data
        If lastRow > 0 Then
            If i = LBound(sheetNames) Then
                ' First sheet: Copy everything including the header
                wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRow, wsSource.Columns.Count)).Copy
            Else
                ' Other sheets: Exclude the header (start from row 2)
                If lastRow > 1 Then
                    wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(lastRow, wsSource.Columns.Count)).Copy
                Else
                    ' If only header exists, skip this sheet
                    GoTo NextSheet
                End If
            End If
            
            ' Paste values into target sheet
            wsTarget.Cells(targetRow, 1).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            
            ' Update next target row
            targetRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row + 1
        End If
        
        NextSheet:
    Next i

    ' MsgBox "Data merged successfully! Now running the primary script.", vbInformation

    ' Call the primary script on the merged data
    Call FilterDataAndCreateSummary

End Sub

