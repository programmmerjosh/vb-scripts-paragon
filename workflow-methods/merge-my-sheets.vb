'=======================
'   Module: MergeTool
'=======================

Sub MergeMySheets()
    Dim wsTarget As Worksheet
    Dim wb As Workbook
    Dim sheetNames As Variant
    Dim i As Integer
    Dim targetRow As Long
    Dim firstSheet As Boolean

    sheetNames = Array("s1", "s2", "s3", "s4", "s5", "s6", "s7", "s8")
    Set wb = ThisWorkbook
    Set wsTarget = GetOrCreateSheet(wb, "special")
    
    targetRow = 1
    firstSheet = True

    For i = LBound(sheetNames) To UBound(sheetNames)
        If SheetExists(sheetNames(i)) Then
            Call CopySheetData(wb.Sheets(sheetNames(i)), wsTarget, targetRow, firstSheet)
            If firstSheet Then firstSheet = False
            targetRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row + 1
        Else
            Debug.Print "Sheet not found: " & sheetNames(i)
        End If
    Next i

    Call FilterDataAndCreateSummary
End Sub

' === Supporting Functions Of MergeMySheets() [START] ===
Function GetOrCreateSheet(wb As Workbook, sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = wb.Sheets(sheetName)
    On Error GoTo 0
    
    If Not GetOrCreateSheet Is Nothing Then
        GetOrCreateSheet.Cells.Clear
    Else
        Set GetOrCreateSheet = wb.Sheets.Add
        GetOrCreateSheet.Name = sheetName
    End If
End Function

Sub CopySheetData(wsSource As Worksheet, wsTarget As Worksheet, ByRef targetRow As Long, isFirstSheet As Boolean)
    Dim lastRow As Long, colCount As Long
    Dim srcRng As Range
    Dim colNum As Long, j As Long
    Dim colData As Variant

    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    colCount = wsSource.UsedRange.Columns.Count
    
    If lastRow = 0 Then Exit Sub ' No data to copy
    
    If isFirstSheet Then
        Set srcRng = wsSource.Range("A1", wsSource.Cells(lastRow, colCount))
    ElseIf lastRow > 1 Then
        Set srcRng = wsSource.Range("A2", wsSource.Cells(lastRow, colCount))
    End If
    
    If srcRng Is Nothing Then Exit Sub

    colNum = Application.Match("WORK_UNIT_CD", wsSource.Rows(1), 0)
    
    If Not IsError(colNum) Then
        wsTarget.Columns(colNum).NumberFormat = "@"
    End If

    wsTarget.Cells(targetRow, 1).Resize(srcRng.Rows.Count, srcRng.Columns.Count).Value = srcRng.Value

    If Not IsError(colNum) Then
        Call FormatWorkUnitColumn(wsTarget, colNum, targetRow + 1, targetRow + srcRng.Rows.Count - 1)
    End If
End Sub

Sub FormatWorkUnitColumn(ws As Worksheet, colNum As Long, startRow As Long, endRow As Long)
    Dim colData As Variant
    Dim i As Long

    ws.Columns(colNum).NumberFormat = "@"
    colData = ws.Range(ws.Cells(startRow, colNum), ws.Cells(endRow, colNum)).Value

    For i = 1 To UBound(colData, 1)
        If IsNumeric(colData(i, 1)) Then
            colData(i, 1) = "'" & colData(i, 1)
        End If
    Next i

    ws.Range(ws.Cells(startRow, colNum), ws.Cells(endRow, colNum)).Value = colData
End Sub
' === Supporting Functions Of MergeMySheets() [END] ===