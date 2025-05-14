
' Working refactored version -> CURRENTLY IN TEST PHASE

Sub FilterDataAndCreateSummary()
    ' Define constants
    Dim cYellow As Long
    Dim cPink As Long
    Dim cRed As Long
    Dim cOrange As Long
    Dim cBlue As Long

    ' Set constants
    cYellow = RGB(255, 255, 0)
    cPink = RGB(238, 143, 204)
    cRed = RGB(255, 111, 145)
    cOrange = RGB(255, 171, 96)
    cBlue = RGB(164, 249, 232)

    Dim wsSpecial As Worksheet
    Dim wsFilteredData As Worksheet
    Dim keepList As Variant

    ' /*
    ' STEP 1: EXPORT DESIRED COLUMNS TO NEW WORKSHEET CALLED FilteredData
    ' */

    Set wsSpecial = GetSheetOrExit("special")
    If wsSpecial Is Nothing Then Exit Sub

    Set wsFilteredData = CreateFilteredDataSheet()
    If wsFilteredData Is Nothing Then Exit Sub

    ' Inital column headings for wsFilteredData:
    ' (1)PARENT_NM, (2)CORP_CD, (3)WORK_UNIT_CD, (4)STMT_CNT, (5)INSERT_CNT, (6)REM_MC_CNT, (7)PLAN_TYPE_CD, (8)MST_MAIL_PROVIDER_CD, 
    ' (9)STD_MAIL_PROVIDER_CD, (10)MAIL_ROUTE
    keepList = Array("PARENT_NM", "CORP_CD", "WORK_UNIT_CD", "STMT_CNT", "INSERT_CNT", "REM_MC_CNT", "PLAN_TYPE_CD", "MST_MAIL_PROVIDER_CD", "STD_MAIL_PROVIDER_CD", "MAIL_ROUTE")

    Call CopyColumnsByHeader(wsSpecial, wsFilteredData, keepList)
    Call SortByColumn(wsFilteredData, "WORK_UNIT_CD")

    ' /*
    ' STEP 2: GET OUTERS BASED ON CORP_CD
    ' */

    Dim wsOutersKey As Worksheet
    Dim datasetCORPCol As Range, planTypeCol As Range, outerCol As Range, workOrderCol As Range
    Dim sortedOuters() As Collection
    Dim lastRowDataset As Long

    Set wsOutersKey = GetSheetOrExit("outerskey")
    If wsOutersKey Is Nothing Then Exit Sub

    ' Validate required columns in each sheet
    If Not ValidateRequiredColumns(wsFilteredData, Array("CORP_CD", "PLAN_TYPE_CD")) Then Exit Sub
    If Not ValidateRequiredColumns(wsOutersKey, Array("CORP_CD", "C5_OUTER", "C4_OUTER", "DL_OUTER")) Then Exit Sub

    Set datasetCORPCol = wsFilteredData.Rows(1).Find("CORP_CD")
    Set workOrderCol = wsFilteredData.Rows(1).Find("WORK_UNIT_CD")
    Set planTypeCol = wsFilteredData.Rows(1).Find("PLAN_TYPE_CD")
    Set outerCol = GetOrCreateColumn(wsFilteredData, "OUTER") ' outer column becomes the last column (11)

    lastRowDataset = GetLastRowBeforeBlanks(wsFilteredData, datasetCORPCol.Column)
    sortedOuters = BuildOuterLookup(wsOutersKey, "CORP_CD")
    Call MapOutersToDataset(wsFilteredData, datasetCORPCol.Column, planTypeCol.Column, outerCol.Column, sortedOuters)

    ' /*
    ' Side STEP: HIGHLIGHT OUTERS WE ALWAYS NEED TO ORDER (even when they have zero inserts)
    ' */

    Call HighlightAlwaysOrderedOuters(wsFilteredData, outerCol.Column, workOrderCol.Column, cOrange)
    Call FormatFilteredDataSheet(wsFilteredData, lastRowDataset)

    ' /*
    ' STEP 3: HIGHLIGHT WORK ORDERS AND INSERTS WHERE INSERTS > 4
    ' */
    
    Call HighlightHighInsertCounts(wsFilteredData, "WORK_UNIT_CD", "INSERT_CNT", 4, RGB(255, 111, 145)) ' cRed

    ' /*
    ' STEP 4: HIGHLIGHT REMAKES
    ' */

    Call HighlightRemakes(wsFilteredData, "REM_MC_CNT", RGB(255, 255, 0)) ' cYellow

    ' /*
    ' STEP 5: CREATE A COLOUR KEY ON FilteredData
    ' */

    Call AddColorKey(wsFilteredData, 1, 6, _
    Array("Remakes", "C4 Outers", "Work Orders Of Jobs With Inserts", "Work Orders Of Jobs With Outers We Should Order (0 Inserts)", "New Entries"), _
    Array(cYellow, cPink, cRed, cOrange, cBlue), _
    "Color Key")

    ' /*
    ' STEP 6: CALCULATE A SUMMARY
    ' */
    Dim summaryStartRow As Long
    Dim summaryEndRow As Long
    Dim summaryData As Variant

    summaryData = GenerateOuterSummary(wsFilteredData, wsOutersKey, lastRowDataset)

    ' COLUMNS:
    ' (1)PARENT_NM, (2)CORP_CD, (3)WORK_UNIT_CD, (4)STMT_CNT, (5)INSERT_CNT, (6)REM_MC_CNT, (7)PLAN_TYPE_CD, (8)MST_MAIL_PROVIDER_CD, 
    ' (9)STD_MAIL_PROVIDER_CD, (10)MAIL_ROUTE, (11)OUTERS

    If Is2DArrayEmpty(summaryData) Then
        wsFilteredData.Cells(lastRowDataset + 14, 1).Value = "No summary data available due to no mapped outers"
        wsFilteredData.Columns(11).Delete ' delete (11)OUTERS because we have no mapped outers in this case
    Else
        summaryStartRow = lastRowDataset + 14
        summaryEndRow = summaryStartRow + UBound(summaryData, 2)
    
        Call WriteSummaryTable(wsFilteredData, summaryData, summaryStartRow)
        Call SortSummary(wsFilteredData, summaryStartRow, summaryEndRow)
        Call FormatSummaryTable(wsFilteredData, summaryStartRow, summaryEndRow)
    End If

    wsFilteredData.Columns(10).Delete ' delete (10)MAIL_ROUTE
    wsFilteredData.Columns(9).Delete ' delete (9)STD_MAIL_PROVIDER_CD
    wsFilteredData.Columns(8).Delete ' delete (8)MST_MAIL_PROVIDER_CD ' BUT, potentially keep this one AND then add logic to highlight "R" for Royal Mail

    ' /*
    ' SIDE (non-essential) STEP: DELETE special WORKSHEET AS WE WILL NO LONGER BE NEEDING IT.
    ' */

    Call DeleteSheetIfExists("special")

    ' /*
    ' SIDE (non-essential) STEP: ADD DATE AND TIME to the right header (for printing purposes).
    ' */

    Call AddTimestampToHeader(wsFilteredData)

    ' /*
    ' STEP 7: HIGHLIGHT NEW ENTRIES (which will only execute if the 'previous' worksheet exists)
    ' */
    Dim arrLatestWorkOrders As Variant, arrPreviousWorkOrders As Variant
    Dim wsPreviousFilteredData As Worksheet
    Dim lastRowPreviousDataset As Long
    Dim prevStmtCol As Range

    If SheetExists("previous") Then
        Set wsPreviousFilteredData = ThisWorkbook.Sheets("previous")
        Set prevStmtCol = wsPreviousFilteredData.Rows(1).Find("STMT_CNT")

        If Not prevStmtCol Is Nothing Then
            lastRowPreviousDataset = GetLastRowBeforeBlanks(wsPreviousFilteredData, prevStmtCol.Column)
        Else
            MsgBox "`STMT_CNT` column not found in 'previous' sheet!", vbExclamation
            Exit Sub
        End If

        arrLatestWorkOrders = GetWorkUnitArray(wsFilteredData, "WORK_UNIT_CD", lastRowDataset)
        arrPreviousWorkOrders = GetWorkUnitArray(wsPreviousFilteredData, "WORK_UNIT_CD", lastRowPreviousDataset)

        Call HighlightNewWorkOrders(wsFilteredData, arrPreviousWorkOrders, arrLatestWorkOrders, "CORP_CD", cBlue)
    Else
        MsgBox "The script has run successfully!!", vbInformation
        MsgBox "Important Note: `previous` worksheet is missing. Rename `FilteredData` to `previous` before you run this script again to see the new entries.", vbInformation
        Exit Sub
    End If

    ' /*
    ' STEP 8: IDENTIFY MISSING VALUES AND APPEND TO SUMMARY
    ' */

    Call AppendMissingWorkUnits(wsFilteredData, arrPreviousWorkOrders, arrLatestWorkOrders, summaryEndRow)

    ' /*
    ' SIDE (non-essential) STEP: DELETE `previous` WORKSHEET AS WE WILL NO LONGER BE NEEDING IT.
    ' */

    Call DeleteSheetIfExists("previous")
    
    MsgBox "The script has run successfully!!", vbInformation

End Sub

' === Supporting Functions Of FilterDataAndCreateSummary() [START] ===

' For STEP 1 (1 of 4)
Function GetSheetOrExit(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheetOrExit = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If GetSheetOrExit Is Nothing Then
        MsgBox "`" & sheetName & "` worksheet is missing.", vbExclamation
    End If
End Function

' For STEP 1 (2 of 4)
Function CreateFilteredDataSheet() As Worksheet
    On Error Resume Next
    Set CreateFilteredDataSheet = ThisWorkbook.Sheets("FilteredData")
    On Error GoTo 0

    If Not CreateFilteredDataSheet Is Nothing Then
        MsgBox "`FilteredData` already exists. Please delete or rename it.", vbExclamation
        Set CreateFilteredDataSheet = Nothing
    Else
        Set CreateFilteredDataSheet = ThisWorkbook.Sheets.Add
        CreateFilteredDataSheet.Name = "FilteredData"
    End If
End Function

' For STEP 1 (3 of 4)
Sub CopyColumnsByHeader(wsSource As Worksheet, wsTarget As Worksheet, columnNames As Variant)
    Dim headers As Range
    Dim colName As Variant
    Dim i As Long, targetCol As Long

    Set headers = wsSource.Rows(1)
    targetCol = 1

    For Each colName In columnNames
        For i = 1 To headers.Columns.count
            If wsSource.Cells(1, i).Value = colName Then
                wsSource.Columns(i).Copy Destination:=wsTarget.Cells(1, targetCol)
                targetCol = targetCol + 1
                Exit For
            End If
        Next i
    Next colName
End Sub

' For STEP 1 (4 of 4)
Sub SortByColumn(ws As Worksheet, columnHeader As String)
    Dim col As Range
    Dim lastRow As Long

    Set col = ws.Rows(1).Find(columnHeader)
    If col Is Nothing Then
        MsgBox "Column '" & columnHeader & "' not found.", vbExclamation
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.count, col.Column).End(xlUp).Row

    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=ws.Columns(col.Column), Order:=xlAscending

    With ws.Sort
        .SetRange ws.UsedRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

' For STEP 2 (1 of 5)
Function GetOrCreateColumn(ws As Worksheet, headerName As String) As Range
    Dim col As Range
    Set col = ws.Rows(1).Find(headerName, LookIn:=xlValues, LookAt:=xlWhole)
    
    If col Is Nothing Then
        Dim lastCol As Long
        lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column + 1
        ws.Cells(1, lastCol).Value = headerName
        Set GetOrCreateColumn = ws.Cells(1, lastCol)
    Else
        Set GetOrCreateColumn = col
    End If
End Function

' For STEP 2 (2 of 5)
Function GetLastRowBeforeBlanks(ws As Worksheet, colIndex As Long, Optional startRow As Long = 2) As Long
    Dim dataCol As Variant
    Dim i As Long
    Dim lastRow As Long

    ' Load the column into an array (from startRow to the bottom of the worksheet)
    With ws
        dataCol = .Range(.Cells(startRow, colIndex), .Cells(.Rows.count, colIndex)).Value
    End With

    ' Scan until we find two consecutive blanks
    For i = 1 To UBound(dataCol) - 1
        If Trim(dataCol(i, 1)) = "" And Trim(dataCol(i + 1, 1)) = "" Then
            Exit For
        ElseIf Trim(dataCol(i, 1)) <> "" Then
            lastRow = i
        End If
    Next i

    ' Return the correct worksheet row number
    GetLastRowBeforeBlanks = lastRow + startRow - 1
End Function

' For STEP 2 (3 of 5)
Function ValidateRequiredColumns(ws As Worksheet, requiredCols As Variant) As Boolean
    Dim colName As Variant
    Dim col As Range

    For Each colName In requiredCols
        Set col = ws.Rows(1).Find(colName, LookIn:=xlValues, LookAt:=xlWhole)
        If col Is Nothing Then
            MsgBox "Missing required column: " & colName, vbExclamation
            ValidateRequiredColumns = False
            Exit Function
        End If
    Next colName
    ValidateRequiredColumns = True
End Function

' For STEP 2 (4 of 5)
Function BuildOuterLookup(ws As Worksheet, corpColName As String) As Collection()
    Dim corpCol As Range, c5Col As Range, c4Col As Range, dlCol As Range
    Dim csWhistlCol As Range, c5RMCol As Range, c5DHLCol As Range, c4RMCol As Range, c4WhistlCol As Range
    Dim lastRow As Long, i As Long, matchLength As Long
    Dim currentCORP As String
    Dim sortedOuters(1 To 6) As Collection ' 1 to 6 meaning the length of the corp??
    

    Set corpCol = ws.Rows(1).Find(corpColName)
    Set c5Col = ws.Rows(1).Find("C5_OUTER")
    Set c4Col = ws.Rows(1).Find("C4_OUTER")
    Set dlCol = ws.Rows(1).Find("DL_OUTER")
    Set csWhistlCol = ws.Rows(1).Find("C5_WHISTL_OUTER")
    Set c5RMCol = ws.Rows(1).Find("C5_RM_OUTER")
    Set c5DHLCol = ws.Rows(1).Find("C5_DHL_OUTER")
    Set c4RMCol = ws.Rows(1).Find("C4_RM_OUTER")
    Set c4WhistlCol = ws.Rows(1).Find("C4_WHISTL_OUTER")

    lastRow = ws.Cells(ws.Rows.count, corpCol.Column).End(xlUp).Row
    For matchLength = 1 To 6
        Set sortedOuters(matchLength) = New Collection
    Next matchLength

    For i = 2 To lastRow
        currentCORP = Trim(ws.Cells(i, corpCol.Column).Value)
        matchLength = Len(currentCORP)
        If matchLength >= 1 And matchLength <= 6 Then
            sortedOuters(matchLength).Add Array(currentCORP, _
                                                ws.Cells(i, c5Col.Column).Value, _
                                                ws.Cells(i, c4Col.Column).Value, _
                                                ws.Cells(i, dlCol.Column).Value, _
                                                ws.Cells(i, csWhistlCol.Column).Value, _ 
                                                ws.Cells(i, c5RMCol.Column).Value, _ 
                                                ws.Cells(i, c5DHLCol.Column).Value, _ 
                                                ws.Cells(i, c4RMCol.Column).Value, _ 
                                                 ws.Cells(i, c4WhistlCol.Column).Value) 
        End If
    Next i

    BuildOuterLookup = sortedOuters
End Function

' For STEP 2 (5 of 5)
Sub MapOutersToDataset(ws As Worksheet, corpCol As Long, planCol As Long, outerCol As Long, sortedOuters() As Collection)
    Dim i As Long, matchLength As Long
    Dim lastRow As Long
    Dim corpVal As String, planVal As String, mappedOuter As String
    Dim mstMailProvider As String, stdMailProvider As String, mailRoute As String
    Dim entry As Variant

    lastRow = ws.Cells(ws.Rows.count, corpCol).End(xlUp).Row

    ' SANTANDER_MAIL_PROVIDER_CODES: T = WHISTL | Y = DHL | R = ROYAL MAIL

    For i = 2 To lastRow
        corpVal = Trim(ws.Cells(i, corpCol).Value)
        planVal = Trim(ws.Cells(i, planCol).Value)
        mstMailProvider = Trim(ws.Cells(i, ws.Rows(1).Find("MST_MAIL_PROVIDER_CD").Column).Value)
        stdMailProvider = Trim(ws.Cells(i, ws.Rows(1).Find("STD_MAIL_PROVIDER_CD").Column).Value)
        mailRoute = Trim(ws.Cells(i, ws.Rows(1).Find("MAIL_ROUTE").Column).Value)

        mappedOuter = ""

        ' Check if PARENT_NM is "Santander Banking"
        If ws.Cells(i, ws.Rows(1).Find("PARENT_NM").Column).Value = "Santander Banking" Then
            ' Map CORP_CD to SB_CORP_CD in outerskey
            For matchLength = 6 To 1 Step -1
                For Each entry In sortedOuters(matchLength)
                    If Left(corpVal, Len(entry(0))) = entry(0) Then
                        ' Determine the OUTER value based on MST_MAIL_PROVIDER_CD, STD_MAIL_PROVIDER_CD, MAIL_ROUTE, and PLAN_TYPE_CD
                        If planVal = "V" Or planVal = "F" Then
                            If mailRoute = "R" Then
                                mappedOuter = entry(7) ' C4_RM_OUTER
                                ws.Cells(i, outerCol).Interior.Color = RGB(238, 143, 204) ' cPink
                            ElseIf mailRoute = "T" Then
                                mappedOuter = entry(8) ' C4_WHISTL_OUTER
                                ws.Cells(i, outerCol).Interior.Color = RGB(238, 143, 204) ' cPink
                            End If
                        Else
                        
                            If mailRoute = "T" Then
                                mappedOuter = entry(4) ' C5_WHISTL_OUTER
                            ElseIf mailRoute = "Y" Then
                                mappedOuter = entry(6) ' C5_DHL_OUTER
                            ElseIf mailRoute = "R" Then
                                mappedOuter = entry(5) ' C5_RM_OUTER
                            End If
                        End If
                        Exit For
                    End If
                Next entry
                If mappedOuter <> "" Then Exit For
            Next matchLength
        Else
            ' Logic for other non-Santander Banking entries
            For matchLength = 6 To 1 Step -1
                For Each entry In sortedOuters(matchLength)
                    If Left(corpVal, Len(entry(0))) = entry(0) Then
                        If planVal = "V" Or planVal = "F" Then
                            mappedOuter = entry(2) ' C4_OUTER
                            ws.Cells(i, outerCol).Interior.Color = RGB(238, 143, 204) ' cPink
                        ElseIf entry(1) <> "" Then
                            mappedOuter = entry(1) ' C5_OUTER
                        Else
                            mappedOuter = entry(3) ' DL_OUTER
                        End If
                        Exit For
                    End If
                Next entry
                If mappedOuter <> "" Then Exit For
            Next matchLength
        End If

        ' Paste the mapped values into the dataset
        ws.Cells(i, outerCol).Value = mappedOuter
    Next i
End Sub

' For SIDE STEP between 2 & 3 (1 of 2)
Sub HighlightAlwaysOrderedOuters(ws As Worksheet, outerColIndex As Long, colNumToHighlight As Long, highlightColor As Variant)
    Dim i As Long, lastRow As Long
    Dim myOuter As String
    Dim outersToOrder As Variant
    Dim count As Integer

    outersToOrder = Array("50023", "BCY03", "BCORPC5AIR", "BARCLPC52", "GCRPC524TNT", "EOP39TNT", "BSMTNT")

    lastRow = ws.Cells(ws.Rows.count, colNumToHighlight).End(xlUp).Row

    For i = 2 To lastRow
        myOuter = ws.Cells(i, outerColIndex).Value
        For count = LBound(outersToOrder) To UBound(outersToOrder)
            If StrComp(outersToOrder(count), myOuter, vbTextCompare) = 0 Then
                ws.Cells(i, colNumToHighlight).Interior.Color = highlightColor
                Exit For
            End If
        Next count
    Next i
End Sub

' For SIDE STEP between 2 & 3 (2 of 2)
Sub FormatFilteredDataSheet(ws As Worksheet, lastRow As Long)
    With ws
        ' Align column H (column 8)
        .Columns(8).HorizontalAlignment = xlLeft

        ' Set column widths (A to H)
        .Columns("A").ColumnWidth = 8
        .Columns("B").ColumnWidth = 8
        .Columns("C").ColumnWidth = 8
        .Columns("D").ColumnWidth = 8.5
        .Columns("E").ColumnWidth = 7.5
        .Columns("F").ColumnWidth = 6.5
        .Columns("G").ColumnWidth = 3
        .Columns("H").ColumnWidth = 16

        ' Apply number formatting with thousand separator (Columns 4, 5, 6)
        .Range(.Cells(2, 4), .Cells(lastRow, 4)).NumberFormat = "#,##0"
        .Range(.Cells(2, 5), .Cells(lastRow, 5)).NumberFormat = "#,##0"
        .Range(.Cells(2, 6), .Cells(lastRow, 6)).NumberFormat = "#,##0"

        ' Apply borders to data
        With .Range(.Cells(1, 1), .Cells(lastRow, 8)).Borders
            .LineStyle = xlContinuous
            .Color = vbBlack
            .Weight = xlThin
        End With
    End With
End Sub

' For STEP 3
Sub HighlightHighInsertCounts(ws As Worksheet, workUnitColName As String, insertColName As String, threshold As Long, colorCode As Long)
    Dim workUnitCol As Range, insertCol As Range
    Dim insertValue As Variant
    Dim lastRow As Long, i As Long

    Set workUnitCol = ws.Rows(1).Find(workUnitColName, LookIn:=xlValues, LookAt:=xlWhole)
    Set insertCol = ws.Rows(1).Find(insertColName, LookIn:=xlValues, LookAt:=xlWhole)

    If workUnitCol Is Nothing Or insertCol Is Nothing Then
        MsgBox "Required columns '" & workUnitColName & "' or '" & insertColName & "' not found!", vbExclamation
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.count, insertCol.Column).End(xlUp).Row

    For i = 2 To lastRow
        insertValue = ws.Cells(i, insertCol.Column).Value
        If IsNumeric(insertValue) And insertValue > threshold Then
            ws.Cells(i, workUnitCol.Column).Interior.Color = colorCode
            ' Optional: also highlight insert count cell itself
            ' ws.Cells(i, insertCol.Column).Interior.Color = colorCode
        End If
    Next i
End Sub

' For STEP 4
Sub HighlightRemakes(ws As Worksheet, remColName As String, highlightColor As Long)
    Dim remCol As Range
    Dim lastRow As Long, lastCol As Long, i As Long

    Set remCol = ws.Rows(1).Find(remColName, LookIn:=xlValues, LookAt:=xlWhole)
    If remCol Is Nothing Then
        MsgBox "`" & remColName & "` column was not found!", vbExclamation
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.count, remCol.Column).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    For i = 2 To lastRow
        If Trim(ws.Cells(i, remCol.Column).Value) <> "" Then
            ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol)).Interior.Color = highlightColor
        End If
    Next i
End Sub

' For STEP 5
Sub AddColorKey(ws As Worksheet, startCol As Long, mergeCols As Long, keyDescriptions As Variant, keyColors As Variant, headingText As String)
    Dim startRow As Long, headingRow As Long, endRow As Long
    Dim i As Long

    ' Find first empty row after data
    startRow = ws.Cells(ws.Rows.count, startCol).End(xlUp).Row + 4
    headingRow = startRow - 1

    ' Add heading row
    With ws.Range(ws.Cells(headingRow, startCol), ws.Cells(headingRow, startCol + mergeCols - 1))
        .Merge
        .Value = headingText
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With

    ' Add key rows
    For i = LBound(keyDescriptions) To UBound(keyDescriptions)
        With ws.Range(ws.Cells(startRow + i, startCol), ws.Cells(startRow + i, startCol + mergeCols - 1))
            .Merge
            .Value = keyDescriptions(i)
            .Interior.Color = keyColors(i)
        End With
    Next i

    ' Apply borders to heading and key
    endRow = ws.Cells(ws.Rows.count, startCol).End(xlUp).Row
    For i = headingRow To endRow
        With ws.Range(ws.Cells(i, startCol), ws.Cells(i, startCol + mergeCols - 1)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
    Next i
End Sub

' For STEP 6 (1 of 5)
Function Is2DArrayEmpty(arr As Variant) As Boolean
    On Error GoTo ErrHandler
    If IsArray(arr) Then
        If Not IsEmpty(arr) Then
            Dim r1 As Long, r2 As Long
            r1 = UBound(arr, 1)
            r2 = UBound(arr, 2)
            Is2DArrayEmpty = False
            Exit Function
        End If
    End If
ErrHandler:
    Is2DArrayEmpty = True
End Function

' For STEP 6 (2 of 5)
Function GenerateOuterSummary(wsFilteredData As Worksheet, wsOutersKey As Worksheet, lastRowDataset As Long) As Variant
    Dim outerArray() As Variant, stmtSumArray() As Double, stockArray() As Variant
    Dim outerValue As String, stmtValue As Double, remMCValue As Variant, planTypeValue As String
    Dim stmtCNCol As Range, remMCCol As Range, planTypeCol As Range, outerCol As Range
    Dim summaryData As Variant
    Dim i As Long, idx As Long, lastRowOutersKey As Long
    Dim stockLocation As String
    Dim foundOuter As Boolean

    ' Find columns
    Set outerCol = wsFilteredData.Rows(1).Find("OUTER")
    Set stmtCNCol = wsFilteredData.Rows(1).Find("STMT_CNT")
    Set remMCCol = wsFilteredData.Rows(1).Find("REM_MC_CNT")
    Set planTypeCol = wsFilteredData.Rows(1).Find("PLAN_TYPE_CD")

    If outerCol Is Nothing Or stmtCNCol Is Nothing Or remMCCol Is Nothing Or planTypeCol Is Nothing Then
        MsgBox "Required columns for summary not found!", vbExclamation
        Exit Function
    End If

    lastRowOutersKey = wsOutersKey.Cells(wsOutersKey.Rows.count, 1).End(xlUp).Row

    ' Initialize arrays
    ReDim outerArray(1 To 1)
    ReDim stmtSumArray(1 To 1)
    ReDim stockArray(1 To 1)

    idx = 0

    For i = 2 To lastRowDataset
        outerValue = Trim(wsFilteredData.Cells(i, outerCol.Column).Value)
        planTypeValue = Trim(wsFilteredData.Cells(i, planTypeCol.Column).Value)
        stmtValue = wsFilteredData.Cells(i, stmtCNCol.Column).Value
        remMCValue = wsFilteredData.Cells(i, remMCCol.Column).Value

        ' Use REM_MC_CNT if it exists
        If Not IsEmpty(remMCValue) And IsNumeric(remMCValue) Then
            stmtValue = remMCValue
        End If

        If outerValue <> "" Then
            foundOuter = False

            ' Check if outer already exists in our list
            For idxCheck = 1 To idx
                If outerArray(idxCheck) = outerValue Then
                    stmtSumArray(idxCheck) = stmtSumArray(idxCheck) + stmtValue
                    foundOuter = True
                    Exit For
                End If
            Next idxCheck

            If Not foundOuter Then
                Dim matched As Boolean: matched = False
                stockLocation = ""

                ' Match in outerskey
                For j = 2 To lastRowOutersKey
                    If planTypeValue = "V" Or planTypeValue = "F" Then
                        If wsOutersKey.Cells(j, 3).Value = outerValue Then
                            stockLocation = wsOutersKey.Cells(j, 6).Value
                            matched = True
                            Exit For
                        End If
                    Else
                        If wsOutersKey.Cells(j, 2).Value = outerValue Then
                            stockLocation = wsOutersKey.Cells(j, 5).Value
                            matched = True
                            Exit For
                        ElseIf wsOutersKey.Cells(j, 4).Value = outerValue Then
                            stockLocation = wsOutersKey.Cells(j, 7).Value
                            matched = True
                            Exit For
                        End If
                    End If
                Next j

                If matched Then
                    idx = idx + 1
                    ReDim Preserve outerArray(1 To idx)
                    ReDim Preserve stmtSumArray(1 To idx)
                    ReDim Preserve stockArray(1 To idx)

                    outerArray(idx) = outerValue
                    stmtSumArray(idx) = stmtValue
                    stockArray(idx) = stockLocation
                End If
            End If
        End If
    Next i

    ' If no data found
    If idx = 0 Then
        GenerateOuterSummary = Array() ' Return empty array
        Exit Function
    End If

    ' Combine into output array
    ReDim summaryData(1 To 3, 1 To idx)
    For i = 1 To idx
        summaryData(1, i) = outerArray(i)
        summaryData(2, i) = stmtSumArray(i)
        summaryData(3, i) = stockArray(i)
    Next i

    GenerateOuterSummary = summaryData
End Function

' For STEP 6 (3 of 5)
Sub WriteSummaryTable(ws As Worksheet, summaryData As Variant, startRow As Long)
    Dim i As Long, rowIdx As Long

    ' Write headers
    ws.Cells(startRow, 1).Value = "OUTER"
    ws.Cells(startRow, 3).Value = "SUM"
    ws.Cells(startRow, 4).Value = "STOCK_LOCATION"

    ' Write summary data row by row
    For i = 1 To UBound(summaryData, 2)
        rowIdx = startRow + i
        ws.Cells(rowIdx, 1).Value = summaryData(1, i) ' OUTER
        ws.Cells(rowIdx, 3).Value = summaryData(2, i) ' SUM
        ws.Cells(rowIdx, 4).Value = summaryData(3, i) ' STOCK_LOCATION
    Next i
End Sub

' For STEP 6 (4 of 5)
    Sub SortSummary(ws As Worksheet, startRow As Long, endRow As Long)
        Dim sortRange As Range
        Set sortRange = ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, 4)) ' Col A to D, unmerged

        With ws.Sort
            .SortFields.Clear
            .SortFields.Add Key:=ws.Columns(1), Order:=xlAscending
            .SetRange sortRange
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End Sub

' For STEP 6 (5 of 5)
Sub FormatSummaryTable(ws As Worksheet, startRow As Long, endRow As Long)
    Dim i As Long
    Dim summaryLastCol As Long: summaryLastCol = 7 ' Merge to col G

    ' Format headers
    With ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, summaryLastCol))
        .Font.Bold = True
        .Font.Italic = True
    End With
    
    With ws
        ' Align column A (column 1)
        .Columns(1).HorizontalAlignment = xlLeft
    End With

    ' Format number column (SUM)
    With ws.Range(ws.Cells(startRow, 3), ws.Cells(endRow, 3))
        .NumberFormat = "#,##0"
    End With

    ' Merge header cells
    ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 2)).Merge
    ws.Range(ws.Cells(startRow, 4), ws.Cells(startRow, summaryLastCol)).Merge

    ' Merge and format each data row
    For i = startRow + 1 To endRow
        ws.Range(ws.Cells(i, 1), ws.Cells(i, 2)).Merge
        ws.Range(ws.Cells(i, 4), ws.Cells(i, summaryLastCol)).Merge
    Next i

    ' Borders
    With ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, summaryLastCol)).Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin
    End With
End Sub

' For SIDE STEP between 6 & 7 (1 of 2) & MergeMySheets()
Sub DeleteSheetIfExists(sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
End Sub

' For SIDE STEP between 6 & 7 (2 of 2)
Sub AddTimestampToHeader(ws As Worksheet)
    Dim formattedDate As String
    formattedDate = Format(Now, "HH:mm - DD/MM/YYYY")

    With ws.PageSetup
        .RightHeader = formattedDate
    End With
End Sub

' For STEP 7 (1 of 2) & MergeMySheets()
Function SheetExists(sheetName As Variant) As Boolean
    On Error Resume Next
    SheetExists = Not ThisWorkbook.Sheets(CStr(sheetName)) Is Nothing
    On Error GoTo 0
End Function

' For STEP 7 & 8 (helper function)
Function GetWorkUnitArray(ws As Worksheet, colName As String, Optional lastRowLimit As Long = 0) As Variant
    Dim colRef As Range
    Set colRef = ws.Rows(1).Find(colName, LookIn:=xlValues, LookAt:=xlWhole)

    If colRef Is Nothing Then
        MsgBox "`" & colName & "` column not found!", vbExclamation
        Exit Function
    End If

    Dim lastRow As Long
    If lastRowLimit > 0 Then
        lastRow = lastRowLimit
    Else
        lastRow = ws.Cells(ws.Rows.count, colRef.Column).End(xlUp).Row
    End If

    GetWorkUnitArray = ws.Range(ws.Cells(2, colRef.Column), ws.Cells(lastRow, colRef.Column)).Value
End Function

' For STEP 7 (2 of 2)
Sub HighlightNewWorkOrders(ws As Worksheet, arrPrevious As Variant, arrLatest As Variant, keyColName As String, highlightColor As Long)
    Dim currentCol As Range
    Dim i As Long, j As Long
    Dim isFound As Boolean

    ' Locate the column in the current worksheet
    Set currentCol = ws.Rows(1).Find(keyColName, LookIn:=xlValues, LookAt:=xlWhole)

    If currentCol Is Nothing Then
        MsgBox "`" & keyColName & "` column not found in the filtered sheet!", vbExclamation
        Exit Sub
    End If

    ' Compare each latest work order to the previous list
    For i = 1 To UBound(arrLatest, 1)
        isFound = False
        For j = 1 To UBound(arrPrevious, 1)
            If arrLatest(i, 1) = arrPrevious(j, 1) Then
                isFound = True
                Exit For
            End If
        Next j

        If Not isFound Then
            ' i + 1 = actual row on worksheet (accounting for header)
            ws.Cells(i + 1, currentCol.Column).Interior.Color = highlightColor
        End If
    Next i
End Sub

' For STEP 8
Sub AppendMissingWorkUnits(ws As Worksheet, arrPrevious As Variant, arrLatest As Variant, summaryEndRow As Long)
    Dim missingValues() As Variant
    Dim missingCount As Long
    Dim i As Long, j As Long
    Dim foundMissing As Boolean
    Dim startRow As Long

    missingCount = 0

    ' Compare each entry in previous against latest
    For i = 1 To UBound(arrPrevious, 1)
        foundMissing = True

        For j = 1 To UBound(arrLatest, 1)
            If arrPrevious(i, 1) = arrLatest(j, 1) Then
                foundMissing = False
                Exit For
            End If
        Next j

        If foundMissing Then
            missingCount = missingCount + 1
            ReDim Preserve missingValues(1 To missingCount)
            missingValues(missingCount) = arrPrevious(i, 1)
        End If
    Next i

    ' Output missing values
    If missingCount > 0 Then
        startRow = summaryEndRow + 2

        ws.Cells(startRow, 1).Value = "ENCLOSED WORK ORDERS"

        For i = 1 To missingCount
            ws.Cells(startRow + i, 1).Value = "'" & missingValues(i)
        Next i

        ' Format the range
        With ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + missingCount, 1))
            .Font.Bold = True
            .Font.Italic = True
            .HorizontalAlignment = xlLeft
            .Borders.LineStyle = xlContinuous
            .Borders.Color = vbBlack
        End With
    End If
End Sub
' === Supporting Functions Of FilterDataAndCreateSummary() [END] ===

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
    Dim anySheet As Boolean

    sheetNames = Array("s1", "s2", "s3", "s4", "s5", "s6", "s7", "s8")
    Set wb = ThisWorkbook
    Set wsTarget = GetOrCreateSheet(wb, "special")
    
    targetRow = 1
    firstSheet = True
    anySheet = False

    For i = LBound(sheetNames) To UBound(sheetNames)
        If SheetExists(sheetNames(i)) Then ' function declared in filter-data-and-create-summary.vb
            anySheet = True
            Call CopySheetData(wb.Sheets(sheetNames(i)), wsTarget, targetRow, firstSheet)
            If firstSheet Then firstSheet = False
            targetRow = wsTarget.Cells(wsTarget.Rows.count, 1).End(xlUp).Row + 1
        Else
            Debug.Print "Sheet not found: " & sheetNames(i)
        End If
    Next i

    If Not anySheet Then
        MsgBox "`" & "s1" & "` worksheet is missing.", vbExclamation
        DeleteSheetIfExists ("special") ' function declared in filter-data-and-create-summary.vb
        Exit Sub
    End If

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

    lastRow = wsSource.Cells(wsSource.Rows.count, 1).End(xlUp).Row
    colCount = wsSource.UsedRange.Columns.count
    
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

    wsTarget.Cells(targetRow, 1).Resize(srcRng.Rows.count, srcRng.Columns.count).Value = srcRng.Value

    If Not IsError(colNum) Then
        Call FormatWorkUnitColumn(wsTarget, colNum, targetRow + 1, targetRow + srcRng.Rows.count - 1)
    End If
End Sub

Sub FormatWorkUnitColumn(ws As Worksheet, startRow As Long, endRow As Long, colNum As Long)
    With ws.Range(ws.Cells(startRow, colNum), ws.Cells(endRow, colNum))
        .NumberFormat = "@" ' Set format to text
        .Value = .Value     ' Force Excel to re-evaluate the values as text
    End With
End Sub
' === Supporting Functions Of MergeMySheets() [END] ===

