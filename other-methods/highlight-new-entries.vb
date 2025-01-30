Sub HighlightNewEntries()
    Dim wsFilteredData As Worksheet, wsPreviousFilteredData As Worksheet
    Dim latestWorkOrderCol As Range, previousWorkOrderCol As Range
    Dim arrLatestWorkOrders As Variant, arrPreviousWorkOrders As Variant
    Dim lastRowFiltered As Long, lastRowPrevious As Long
    Dim lastColFiltered As Long, lastColPrevious As Long
    Dim i As Long, j As Long
    Dim isFound As Boolean
    Dim cell As Range

    ' Set the target worksheets
    On Error Resume Next
    Set wsFilteredData = ThisWorkbook.Sheets("FilteredData")
    Set wsPreviousFilteredData = ThisWorkbook.Sheets("previous")
    On Error GoTo 0
    
     ' Check if wsFilteredData exists
    If wsFilteredData Is Nothing Then
        MsgBox "FilteredData worksheet is missing!", vbExclamation
        Exit Sub
    End If

    ' If wsPreviousFilteredData is missing, just skip comparison
    If wsPreviousFilteredData Is Nothing Then
        MsgBox "previous worksheet is missing. Skipping comparison.", vbInformation
        Exit Sub
    End If

    ' Find the last row and column in both worksheets
    lastRowFiltered = wsFilteredData.Cells(wsFilteredData.Rows.Count, "A").End(xlUp).Row
    lastRowPrevious = wsPreviousFilteredData.Cells(wsPreviousFilteredData.Rows.Count, "A").End(xlUp).Row
    lastColFiltered = wsFilteredData.Cells(1, wsFilteredData.Columns.Count).End(xlToLeft).Column
    lastColPrevious = wsPreviousFilteredData.Cells(1, wsPreviousFilteredData.Columns.Count).End(xlToLeft).Column

    ' Identify the "WORK_UNIT_CD" column in both sheets (assuming column name is in the header)
    Set latestWorkOrderCol = wsFilteredData.Rows(1).Find("WORK_UNIT_CD", LookIn:=xlValues, LookAt:=xlWhole)
    Set previousWorkOrderCol = wsPreviousFilteredData.Rows(1).Find("WORK_UNIT_CD", LookIn:=xlValues, LookAt:=xlWhole)

    ' Check if "WORK_UNIT_CD" columns are found
    If latestWorkOrderCol Is Nothing Or previousWorkOrderCol Is Nothing Then
        MsgBox "WORK_UNIT_CD column not found!", vbExclamation
        Exit Sub
    End If

    ' Load WORK_UNIT_CD values into arrays
    arrLatestWorkOrders = wsFilteredData.Range(wsFilteredData.Cells(2, latestWorkOrderCol.Column), wsFilteredData.Cells(lastRowFiltered, latestWorkOrderCol.Column)).value
    arrPreviousWorkOrders = wsPreviousFilteredData.Range(wsPreviousFilteredData.Cells(2, previousWorkOrderCol.Column), wsPreviousFilteredData.Cells(lastRowPrevious, previousWorkOrderCol.Column)).value

    ' Compare the arrays and highlight rows where WORK_UNIT_CD in arrLatestWorkOrders is not found in arrPreviousWorkOrders
    For i = 1 To UBound(arrLatestWorkOrders, 1)
        isFound = False

        ' Compare each work order in arrLatestWorkOrders with arrPreviousWorkOrders
        For j = 1 To UBound(arrPreviousWorkOrders, 1)
            If arrLatestWorkOrders(i, 1) = arrPreviousWorkOrders(j, 1) Then
                isFound = True
                Exit For
            End If
        Next j

        ' If the work order is not found, highlight the row in green
        If Not isFound Then
            wsFilteredData.Range(wsFilteredData.Cells(i + 1, 1), wsFilteredData.Cells(i + 1, lastColFiltered)).Interior.Color = RGB(164, 249, 232) ' baby blue

        End If
    Next i

    ' MsgBox "Rows with new work orders have been highlighted.", vbInformation
End Sub


