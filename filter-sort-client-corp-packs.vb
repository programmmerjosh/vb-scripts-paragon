' This script does the following:
'   1. Filters out unwanted columns
'   2. Primary sort: PARENT_NM
'   3. Secondary sort: CORP_CD (first 3 characters)
'   4. Tertiary sort: STMT_CNT
'   5. Takes into account that the STMT_CNT is irrelevent if it has a remake (REM_MC_CNT) value, so the sort uses REM_MC_CNT instead of STMT_CNT

'IMPORTANT NOTE: Don't forget to change the sheet name to, "Special1" before running this Macro

Sub FilterSortClientCorpPacks()
    Dim ws As Worksheet
    Dim colName As Variant
    Dim headers As Range
    Dim keepList As Variant
    Dim deleteColumn As Boolean
    Dim i As Long
    Dim corpIDCol As Range, stmtCNCol As Range, parantNMCol As Range, remMCCol As Range, helperCol As Range
    Dim lastRow As Long
    Dim helperColIndex As Long, stmtHelperColIndex As Long

    ' Set worksheet
    Set ws = ThisWorkbook.Sheets("Special1") ' Update to your sheet name

    ' Step 1: Keep only the specified columns
    keepList = Array("PARENT_NM", "CORP_CD", "WORK_UNIT_CD", "STMT_CNT", "INSERT_CNT", "REM_MC_CNT", "PLAN_TYPE_CD") ' Desired column names
    Set headers = ws.Rows(1) ' Assuming headers are in row 1

    ' Delete columns not in keepList
    For i = ws.Columns.Count To 1 Step -1
        deleteColumn = True
        For Each colName In keepList
            If ws.Cells(1, i).Value = colName Then
                deleteColumn = False
                Exit For
            End If
        Next colName
        If deleteColumn Then ws.Columns(i).Delete
    Next i

    ' Step 2: Find relevant columns
    Set parantNMCol = ws.Rows(1).Find("PARENT_NM")
    Set corpIDCol = ws.Rows(1).Find("CORP_CD")
    Set stmtCNCol = ws.Rows(1).Find("STMT_CNT")
    Set remMCCol = ws.Rows(1).Find("REM_MC_CNT")

    ' Validate that columns exist
    If parantNMCol Is Nothing Or corpIDCol Is Nothing Or stmtCNCol Is Nothing Or remMCCol Is Nothing Then
        MsgBox "Required columns (PARENT_NM, CORP_CD, STMT_CNT, REM_MC_CNT) not found!", vbExclamation
        Exit Sub
    End If

    ' Step 3: Add a helper column for the first 3 characters of CORP_CD
    lastRow = ws.Cells(ws.Rows.Count, corpIDCol.Column).End(xlUp).Row
    helperColIndex = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
    ws.Cells(1, helperColIndex).Value = "Helper_CORP_CD"

    Dim j As Long
    For j = 2 To lastRow
        ws.Cells(j, helperColIndex).Value = Left(ws.Cells(j, corpIDCol.Column).Value, 3)
    Next j
    Set helperCol = ws.Columns(helperColIndex)

    ' Step 4: Add a helper column for the adjusted STMT_CNT values
    stmtHelperColIndex = helperColIndex + 1
    ws.Cells(1, stmtHelperColIndex).Value = "Adjusted_STMT_CNT"

    For j = 2 To lastRow
        If ws.Cells(j, remMCCol.Column).Value <> "" Then
            ws.Cells(j, stmtHelperColIndex).Value = ws.Cells(j, remMCCol.Column).Value
        Else
            ws.Cells(j, stmtHelperColIndex).Value = ws.Cells(j, stmtCNCol.Column).Value
        End If
    Next j
    Dim stmtHelperCol As Range
    Set stmtHelperCol = ws.Columns(stmtHelperColIndex)

    ' Step 5: Sort data
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=ws.Columns(parantNMCol.Column), Order:=xlAscending ' Primary: PARENT_NM
    ws.Sort.SortFields.Add Key:=helperCol, Order:=xlAscending ' Secondary: Helper_CORP_CD
    ws.Sort.SortFields.Add Key:=stmtHelperCol, Order:=xlDescending ' Tertiary: Adjusted_STMT_CNT

    With ws.Sort
        .SetRange ws.UsedRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    ' Step 6: Remove helper columns
    helperCol.Delete
    stmtHelperCol.Delete
End Sub
