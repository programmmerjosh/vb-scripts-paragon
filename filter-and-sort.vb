Sub FilterAndSort()

    ' /*
    ' STEP 1: EXPORT DESIRED COLUMNS TO NEW WORKSHEET CALLED FilteredData
    ' */

    Dim wsSpecial As Worksheet, wsFilteredData As Worksheet
    Dim colName As Variant
    Dim headers As Range
    Dim keepList As Variant
    Dim i As Long
    Dim stmtCnt As Range, parantNMCol As Range, remMCCol As Range, stmtHelperCol As Range
    Dim lastRow As Long
    Dim targetCol As Long
    Dim helperColIndex As Long

    ' Set source worksheet
    Set wsSpecial = ThisWorkbook.Sheets("Special") ' Update to your sheet name

    ' Create a new worksheet for the filtered data
    Set wsFilteredData = ThisWorkbook.Sheets.Add
    wsFilteredData.Name = "FilteredData" ' Change as needed

    ' Define the columns to keep
    keepList = Array("PARENT_NM", "CORP_CD", "WORK_UNIT_CD", "STMT_CNT", "INSERT_CNT", "REM_MC_CNT", "PLAN_TYPE_CD") ' Desired column names
    Set headers = wsSpecial.Rows(1) ' Assuming headers are in row 1

    ' Step 1: Copy the keepList columns to the new worksheet
    targetCol = 1
    For Each colName In keepList
        For i = 1 To headers.Columns.Count
            If wsSpecial.Cells(1, i).Value = colName Then
                wsSpecial.Columns(i).Copy Destination:=wsFilteredData.Cells(1, targetCol)
                targetCol = targetCol + 1
                Exit For
            End If
        Next i
    Next colName

    ' Step 2: Set reference for columns
    Set stmtCnt = wsFilteredData.Rows(1).Find("STMT_CNT")
    Set parantNMCol = wsFilteredData.Rows(1).Find("PARENT_NM")
    Set remMCCol = wsFilteredData.Rows(1).Find("REM_MC_CNT")
    Set remMCCol = wsFilteredData.Rows(1).Find("REM_MC_CNT")

    ' Validate that columns exist
    If parantNMCol Is Nothing Or stmtCnt Is Nothing Or remMCCol Is Nothing Then
        MsgBox "Required columns not found!", vbExclamation
        Exit Sub
    End If

    ' Step 3: Add a helper column for the adjusted STMT_CNT values
    helperColIndex = wsFilteredData.Cells(1, wsFilteredData.Columns.Count).End(xlToLeft).Column + 1
    wsFilteredData.Cells(1, helperColIndex).Value = "Adjusted_STMT_CNT"

    lastRow = wsFilteredData.Cells(wsFilteredData.Rows.Count, parantNMCol.Column).End(xlUp).Row

    Dim j As Long
    For j = 2 To lastRow
        If wsFilteredData.Cells(j, remMCCol.Column).Value <> "" Then
            wsFilteredData.Cells(j, helperColIndex).Value = wsFilteredData.Cells(j, remMCCol.Column).Value
        Else
            wsFilteredData.Cells(j, helperColIndex).Value = wsFilteredData.Cells(j, stmtCnt.Column).Value
        End If
    Next j

    ' Step 4: Sort data
    Dim sortRange As Range
    Set sortRange = wsFilteredData.Range(wsFilteredData.Cells(1, 1), wsFilteredData.Cells(lastRow, helperColIndex))

    wsFilteredData.Sort.SortFields.Clear
    wsFilteredData.Sort.SortFields.Add Key:=wsFilteredData.Columns(parantNMCol.Column), Order:=xlAscending ' Primary: PARENT_NM
    wsFilteredData.Sort.SortFields.Add Key:=wsFilteredData.Columns(helperColIndex), Order:=xlDescending ' Secondary: Adjusted_STMT_CNT

    With wsFilteredData.Sort
        .SetRange sortRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    ' Clean up: remove helper column
    Set stmtHelperCol = wsFilteredData.Rows(1).Find("Adjusted_STMT_CNT")
    stmtHelperCol.Delete

    ' Step 5: Apply formatting
    With wsFilteredData
        ' Left-align first column
        .Columns(1).HorizontalAlignment = xlLeft

        ' Shrink column widths where necessary
        Columns("A").ColumnWidth = 12
        Columns("B").ColumnWidth = 8
        Columns("C").ColumnWidth = 8
        Columns("D").ColumnWidth = 6.5
        Columns("E").ColumnWidth = 5
        Columns("F").ColumnWidth = 4
        Columns("G").ColumnWidth = 3
    End With

    ' /*
    ' STEP 2: HIGHLIGHT REMAKES
    ' */
    Dim remCountCol As Range
    Dim lastCol As Long

    ' Find the REM_MC_CNT column
    Set remCountCol = wsFilteredData.Rows(1).Find("REM_MC_CNT")

    ' Validate that REM_MC_CNT column exists
    If remCountCol Is Nothing Then
        MsgBox "The REM_MC_CNT column was not found!", vbExclamation
        Exit Sub
    End If

    ' Find the last row in the REM_MC_CNT column
    lastRow = wsFilteredData.Cells(wsFilteredData.Rows.Count, remCountCol.Column).End(xlUp).Row

    ' Find the last column in the header row
    lastCol = wsFilteredData.Cells(1, wsFilteredData.Columns.Count).End(xlToLeft).Column

    ' Loop through each row in the REM_MC_CNT column
    For i = 2 To lastRow ' Assuming headers are in row 1
        If wsFilteredData.Cells(i, remCountCol.Column).Value <> "" Then
            ' Highlight the row up to the last column
            wsFilteredData.Range(wsFilteredData.Cells(i, 1), wsFilteredData.Cells(i, lastCol)).Interior.Color = RGB(255, 255, 0) ' Yellow color
        End If
    Next i

End Sub
