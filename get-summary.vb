'IMPORTANT NOTE: Don't forget to change the sheet name to, "Special1" before running this Macro

Sub GetSummary()
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

    Dim remCountCol As Range
    Dim lastCol As Long

    ' Find the REM_MC_CNT column
    Set remCountCol = ws.Rows(1).Find("REM_MC_CNT")

    ' Validate that REM_MC_CNT column exists
    If remCountCol Is Nothing Then
        MsgBox "The REM_MC_CNT column was not found!", vbExclamation
        Exit Sub
    End If

    ' Find the last row in the REM_MC_CNT column
    lastRow = ws.Cells(ws.Rows.Count, remCountCol.Column).End(xlUp).Row

    ' Find the last column in the header row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Loop through each row in the REM_MC_CNT column
    For i = 2 To lastRow ' Assuming headers are in row 1
        If ws.Cells(i, remCountCol.Column).Value <> "" Then
            ' Highlight the row up to the last column
            ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol)).Interior.Color = RGB(255, 255, 0) ' Yellow color
        End If
    Next i

    
    Dim workUnitCol As Range, insertCntCol As Range
    Dim insertCntValue As Variant
    
    ' Find the columns for WORK_UNIT_CD and INSERT_CNT
    Set workUnitCol = ws.Rows(1).Find("WORK_UNIT_CD")
    Set insertCntCol = ws.Rows(1).Find("INSERT_CNT")
    
    ' Validate the columns exist
    If workUnitCol Is Nothing Or insertCntCol Is Nothing Then
        MsgBox "Required columns (WORK_UNIT_CD, INSERT_CNT) not found!", vbExclamation
        Exit Sub
    End If
    
    ' Find the last row of data
    lastRow = ws.Cells(ws.Rows.Count, workUnitCol.Column).End(xlUp).Row
    
    ' Loop through each row to check the INSERT_CNT value
    For i = 2 To lastRow
        insertCntValue = ws.Cells(i, insertCntCol.Column).Value
        
        ' Check if INSERT_CNT is greater than 9
        If IsNumeric(insertCntValue) And insertCntValue > 9 Then
            ' Highlight WORK_UNIT_CD with rgb(255,111,145)
            ws.Cells(i, workUnitCol.Column).Interior.Color = RGB(255, 111, 145)
            ' Highlight INSERT_CNT with rgb(255,171,96)
            ws.Cells(i, insertCntCol.Column).Interior.Color = RGB(255, 171, 96)
        End If
    Next i

    
    Dim wsOutersKey As Worksheet
    Dim datasetCORPCol As Range, outerCol As Range
    Dim outC5OuterCol As Range, outC4OuterCol As Range, outDLOuterCol As Range
    Dim planTypeCol As Range
    Dim lastRowDataset As Long, lastRowOutersKey As Long
    Dim matchLength As Long
    Dim datasetCORPValue As String, planTypeValue As String, mappedOuter As String
    Dim sortedOuters(1 To 6) As Collection ' Array of collections to group by length
    Dim entry As Variant

    ' Set the OUTERSKEY worksheets
    Set wsOutersKey = ThisWorkbook.Sheets("OUTERSKEY") ' Update to your OUTERSKEY sheet name

    ' Find the relevant columns in the dataset
    Set datasetCORPCol = ws.Rows(1).Find("CORP_CD")
    Set planTypeCol = ws.Rows(1).Find("PLAN_TYPE_CD")

    ' Check if OUTER column exists
    Set outerCol = ws.Rows(1).Find("OUTER")

    ' If OUTER column doesn't exist, add it at the next empty column
    If outerCol Is Nothing Then
        Dim lastColumn As Long
        lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
        ws.Cells(1, lastColumn).Value = "OUTER"
        Set outerCol = ws.Cells(1, lastColumn) ' Now set outerCol to the newly created column
    End If

    ' Find the relevant columns in OUTERSKEY
    Set outC5OuterCol = wsOutersKey.Rows(1).Find("C5_OUTER")
    Set outC4OuterCol = wsOutersKey.Rows(1).Find("C4_OUTER")
    Set outDLOuterCol = wsOutersKey.Rows(1).Find("DL_OUTER")
    Set outerCORPCol = wsOutersKey.Rows(1).Find("CORP_CD")

    ' Validate columns exist
    If datasetCORPCol Is Nothing Or planTypeCol Is Nothing Then
        MsgBox "Required columns (CORP_CD, PLAN_TYPE_CD) not found in the dataset!", vbExclamation
        Exit Sub
    End If
    If outerCORPCol Is Nothing Or outC5OuterCol Is Nothing Or outC4OuterCol Is Nothing Or outDLOuterCol Is Nothing Then
        MsgBox "Required columns (CORP_CD, C5_OUTER, C4_OUTER, DL_OUTER) not found in OUTERSKEY!", vbExclamation
        Exit Sub
    End If

    ' Initialize the sortedOuters array
    For matchLength = 1 To 6
        Set sortedOuters(matchLength) = New Collection
    Next matchLength

    ' Populate sortedOuters based on CORP_CD length
    lastRowOutersKey = wsOutersKey.Cells(wsOutersKey.Rows.Count, outerCORPCol.Column).End(xlUp).Row
    For i = 2 To lastRowOutersKey
        Dim currentCORP As String
        currentCORP = Trim(wsOutersKey.Cells(i, outerCORPCol.Column).Value)
        matchLength = Len(currentCORP)
        If matchLength >= 1 And matchLength <= 6 Then
            sortedOuters(matchLength).Add Array(currentCORP, _
                                                wsOutersKey.Cells(i, outC5OuterCol.Column).Value, _
                                                wsOutersKey.Cells(i, outC4OuterCol.Column).Value, _
                                                wsOutersKey.Cells(i, outDLOuterCol.Column).Value)
        End If
    Next i

    ' Map the OUTERSKEY entries to the dataset
    lastRowDataset = ws.Cells(ws.Rows.Count, datasetCORPCol.Column).End(xlUp).Row
    For i = 2 To lastRowDataset
        datasetCORPValue = Trim(ws.Cells(i, datasetCORPCol.Column).Value)
        planTypeValue = Trim(ws.Cells(i, planTypeCol.Column).Value)
        mappedOuter = ""

        ' Check sortedOuters starting from the longest CORP_CD
        For matchLength = 6 To 1 Step -1
            For Each entry In sortedOuters(matchLength)
                If Left(datasetCORPValue, Len(entry(0))) = entry(0) Then
                    If planTypeValue = "U" Or planTypeValue = "F" Then
                        mappedOuter = entry(2) ' Use C4_OUTER
                    ElseIf entry(1) <> "" Then
                        mappedOuter = entry(1) ' Use C5_OUTER
                    Else
                        mappedOuter = entry(3) ' Use DL_OUTER
                    End If
                    Exit For
                End If
            Next entry
            If mappedOuter <> "" Then Exit For
        Next matchLength

        ' Update the OUTER column
        ws.Cells(i, outerCol.Column).Value = mappedOuter
    Next i

    ' Convert OUTER column to text format
    For i = 2 To lastRowDataset
        ws.Cells(i, outerCol.Column).Value = CStr(ws.Cells(i, outerCol.Column).Value) ' Convert to string
    Next i

    ' Apply Text format to the entire OUTER column
    ws.Columns(outerCol.Column).NumberFormat = "@"

    
    Dim wsSummary As Worksheet
    Dim outerValue As String
    Dim stmtValue As Double, remMCValue As Variant
    Dim summaryData As Collection
    Dim key As Variant
    Dim summaryRow As Long
    Dim outerArray() As Variant
    Dim stmtSumArray() As Double
    Dim foundOuter As Boolean
    Dim idx As Long
    
    ' Find the columns for OUTER, STMT_CNT, and REM_MC_CNT
    Set outerCol = ws.Rows(1).Find("OUTER")
    Set stmtCNCol = ws.Rows(1).Find("STMT_CNT")
    Set remMCCol = ws.Rows(1).Find("REM_MC_CNT") ' Column for REM_MC_CNT

    ' Validate the columns exist
    If outerCol Is Nothing Or stmtCNCol Is Nothing Or remMCCol Is Nothing Then
        MsgBox "Required columns (OUTER, STMT_CNT, REM_MC_CNT) not found!", vbExclamation
        Exit Sub
    End If

    ' Find the last row of data
    lastRow = ws.Cells(ws.Rows.Count, outerCol.Column).End(xlUp).Row
    
    ' Initialize the arrays for Outer values and its sum
    ReDim outerArray(1 To 1) ' Initially size to 1 element
    ReDim stmtSumArray(1 To 1) ' Initially size to 1 element
    
    ' Loop through each row to calculate the sum of OUTER
    For i = 2 To lastRow
        outerValue = ws.Cells(i, outerCol.Column).Value
        stmtValue = ws.Cells(i, stmtCNCol.Column).Value
        remMCValue = ws.Cells(i, remMCCol.Column).Value ' Get REM_MC_CNT value

        ' If REM_MC_CNT has a value, use it instead of STMT_CNT
        If Not IsEmpty(remMCValue) And IsNumeric(remMCValue) Then
            stmtValue = remMCValue
        End If

        If outerValue <> "" Then
            foundOuter = False
            ' Check if OUTER value already exists in the array
            For idx = 1 To UBound(outerArray)
                If outerArray(idx) = outerValue Then
                    stmtSumArray(idx) = stmtSumArray(idx) + stmtValue
                    foundOuter = True
                    Exit For
                End If
            Next idx

            ' If OUTER value not found, add new entry
            If Not foundOuter Then
                ReDim Preserve outerArray(1 To UBound(outerArray) + 1)
                ReDim Preserve stmtSumArray(1 To UBound(stmtSumArray) + 1)
                
                outerArray(UBound(outerArray)) = outerValue
                stmtSumArray(UBound(stmtSumArray)) = stmtValue
            End If
        End If
    Next i

    ' Create a summary worksheet
    On Error Resume Next
    Set wsSummary = ThisWorkbook.Sheets("Summary")
    If wsSummary Is Nothing Then
        Set wsSummary = ThisWorkbook.Sheets.Add
        wsSummary.Name = "Summary"
    Else
        wsSummary.Cells.Clear
    End If
    On Error GoTo 0

    ' Write the summary to the new worksheet
    wsSummary.Cells(1, 1).Value = "OUTER"
    wsSummary.Cells(1, 2).Value = "SUM"

    ' Write the data into the summary sheet
    summaryRow = 2
    For idx = 1 To UBound(outerArray)
        wsSummary.Cells(summaryRow, 1).Value = outerArray(idx)
        wsSummary.Cells(summaryRow, 2).Value = stmtSumArray(idx)
        summaryRow = summaryRow + 1
    Next idx

    ' Cleanup: Delete rows where OUTER is blank
    Dim rowIdx As Long
    rowIdx = summaryRow - 1 ' Last row with data in summary
    For rowIdx = rowIdx To 2 Step -1
        If wsSummary.Cells(rowIdx, 1).Value = "" Then
            wsSummary.Rows(rowIdx).Delete
        End If
    Next rowIdx
End Sub
