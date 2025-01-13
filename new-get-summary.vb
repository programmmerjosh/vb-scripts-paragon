Sub NewGetSummary()

' /*
' STEP 1: EXPORT DESIRED COLUMNS TO NEW WORKSHEET CALLED SUMMARY
' */

    Dim wsSpecial As Worksheet, wsFilteredData As Worksheet
    Dim colName As Variant
    Dim headers As Range
    Dim keepList As Variant
    Dim i As Long
    Dim workUnitCol As Range
    Dim lastRow As Long
    Dim targetCol As Long

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

    ' Step 2: Set reference for WORK_UNIT_CD column
    Set workUnitCol = wsFilteredData.Rows(1).Find("WORK_UNIT_CD")

    ' Validate that WORK_UNIT_CD column exists
    If workUnitCol Is Nothing Then
        MsgBox "Required column (WORK_UNIT_CD) not found!", vbExclamation
        Exit Sub
    End If

    ' Step 3: Sort data by WORK_UNIT_CD
    lastRow = wsFilteredData.Cells(wsFilteredData.Rows.Count, workUnitCol.Column).End(xlUp).Row
    wsFilteredData.Sort.SortFields.Clear
    wsFilteredData.Sort.SortFields.Add Key:=wsFilteredData.Columns(workUnitCol.Column), Order:=xlAscending ' Sort by WORK_UNIT_CD

    With wsFilteredData.Sort
        .SetRange wsFilteredData.UsedRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

' /*
' STEP 2: GET OUTERS
' */
    Dim wsOutersKey As Worksheet
    Dim datasetCORPCol As Range, outerCol As Range, stockLocationCol As Range
    Dim outC5OuterCol As Range, outC4OuterCol As Range, outDLOuterCol As Range
    Dim outC5StockCol As Range, outC4StockCol As Range, outDLStockCol As Range
    Dim planTypeCol As Range
    Dim lastRowDataset As Long, lastRowOutersKey As Long
    Dim matchLength As Long
    Dim datasetCORPValue As String, planTypeValue As String, mappedOuter As String, mappedStock As String
    Dim sortedOuters(1 To 6) As Collection ' Array of collections to group by length
    Dim entry As Variant

    ' Set the dataset and OUTERSKEY worksheets
    Set wsFilteredData = ThisWorkbook.Sheets("FilteredData") ' Update to your dataset sheet name
    Set wsOutersKey = ThisWorkbook.Sheets("OUTERSKEY") ' Update to your OUTERSKEY sheet name

    ' Find the relevant columns in the dataset
    Set datasetCORPCol = wsFilteredData.Rows(1).Find("CORP_CD")
    Set planTypeCol = wsFilteredData.Rows(1).Find("PLAN_TYPE_CD")

    ' Check if OUTER column exists, add it if not
    Set outerCol = wsFilteredData.Rows(1).Find("OUTER")
    If outerCol Is Nothing Then
        Dim lastColumn As Long
        lastColumn = wsFilteredData.Cells(1, wsFilteredData.Columns.Count).End(xlToLeft).Column + 1
        wsFilteredData.Cells(1, lastColumn).value = "OUTER"
        Set outerCol = wsFilteredData.Cells(1, lastColumn)
    End If

    ' Check if STOCK_LOCATION column exists, add it if not
    Set stockLocationCol = wsFilteredData.Rows(1).Find("STOCK_LOCATION")
    If stockLocationCol Is Nothing Then
        lastColumn = wsFilteredData.Cells(1, wsFilteredData.Columns.Count).End(xlToLeft).Column + 1
        wsFilteredData.Cells(1, lastColumn).value = "STOCK_LOCATION"
        Set stockLocationCol = wsFilteredData.Cells(1, lastColumn)
    End If

    ' Find the relevant columns in OUTERSKEY
    Set outC5OuterCol = wsOutersKey.Rows(1).Find("C5_OUTER")
    Set outC4OuterCol = wsOutersKey.Rows(1).Find("C4_OUTER")
    Set outDLOuterCol = wsOutersKey.Rows(1).Find("DL_OUTER")
    Set outC5StockCol = wsOutersKey.Rows(1).Find("C5_STOCK_LOCATION")
    Set outC4StockCol = wsOutersKey.Rows(1).Find("C4_STOCK_LOCATION")
    Set outDLStockCol = wsOutersKey.Rows(1).Find("DL_STOCK_LOCATION")
    Set outerCORPCol = wsOutersKey.Rows(1).Find("CORP_CD")

    ' Validate columns exist
    If datasetCORPCol Is Nothing Or planTypeCol Is Nothing Then
        MsgBox "Required columns (CORP_CD, PLAN_TYPE_CD) not found in the dataset!", vbExclamation
        Exit Sub
    End If
    If outerCORPCol Is Nothing Or outC5OuterCol Is Nothing Or outC4OuterCol Is Nothing Or outDLOuterCol Is Nothing Or _
       outC5StockCol Is Nothing Or outC4StockCol Is Nothing Or outDLStockCol Is Nothing Then
        MsgBox "Required columns not found in OUTERSKEY!", vbExclamation
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
        currentCORP = Trim(wsOutersKey.Cells(i, outerCORPCol.Column).value)
        matchLength = Len(currentCORP)
        If matchLength >= 1 And matchLength <= 6 Then
            sortedOuters(matchLength).Add Array(currentCORP, _
                                                wsOutersKey.Cells(i, outC5OuterCol.Column).value, _
                                                wsOutersKey.Cells(i, outC4OuterCol.Column).value, _
                                                wsOutersKey.Cells(i, outDLOuterCol.Column).value, _
                                                wsOutersKey.Cells(i, outC5StockCol.Column).value, _
                                                wsOutersKey.Cells(i, outC4StockCol.Column).value, _
                                                wsOutersKey.Cells(i, outDLStockCol.Column).value)
        End If
    Next i

    ' Map the OUTERSKEY entries to the dataset
    lastRowDataset = wsFilteredData.Cells(wsFilteredData.Rows.Count, datasetCORPCol.Column).End(xlUp).Row
    For i = 2 To lastRowDataset
        datasetCORPValue = Trim(wsFilteredData.Cells(i, datasetCORPCol.Column).value)
        planTypeValue = Trim(wsFilteredData.Cells(i, planTypeCol.Column).value)
        mappedOuter = ""
        mappedStock = ""

        ' Check sortedOuters starting from the longest CORP_CD
        For matchLength = 6 To 1 Step -1
            For Each entry In sortedOuters(matchLength)
                If Left(datasetCORPValue, Len(entry(0))) = entry(0) Then
                    If planTypeValue = "V" Or planTypeValue = "F" Then
                        mappedOuter = entry(2) ' Use C4_OUTER
                        mappedStock = entry(5) ' Use C4_STOCK_LOCATION
                    ElseIf entry(1) <> "" Then
                        mappedOuter = entry(1) ' Use C5_OUTER
                        mappedStock = entry(4) ' Use C5_STOCK_LOCATION
                    Else
                        mappedOuter = entry(3) ' Use DL_OUTER
                        mappedStock = entry(6) ' Use DL_STOCK_LOCATION
                    End If
                    Exit For
                End If
            Next entry
            If mappedOuter <> "" Then Exit For
        Next matchLength

        ' Update the OUTER and STOCK_LOCATION columns
        wsFilteredData.Cells(i, outerCol.Column).value = mappedOuter
        wsFilteredData.Cells(i, stockLocationCol.Column).value = mappedStock
    Next i

    ' Convert columns to text format
    wsFilteredData.Columns(outerCol.Column).NumberFormat = "@"
    wsFilteredData.Columns(stockLocationCol.Column).NumberFormat = "@"
    
End Sub
