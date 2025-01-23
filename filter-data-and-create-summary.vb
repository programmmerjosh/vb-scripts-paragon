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

' /*
' STEP 1: REMOVE DUPLICATE HEADING ROWS
' */

    Dim wsSpecial As Worksheet, wsFilteredData As Worksheet

    ' Set source worksheet
    On Error Resume Next
    Set wsSpecial = ThisWorkbook.Sheets("special") ' Update to your sheet name
    On Error GoTo 0
    

    ' If wsSpecial is missing, just exit
    If wsSpecial Is Nothing Then
        MsgBox "`special` worksheet is missing. Please rename your worksheet to `special`", vbInformation
        Exit Sub
    End If

    Dim searchRange As Range, foundCell As Range, deleteRows As Range
    Dim firstAddress As String

    ' Define the initial range to search in (first column, excluding the header row)
    Set searchRange = wsSpecial.Range(wsSpecial.Cells(2, 1), wsSpecial.Cells(wsSpecial.Rows.Count, 1))

    ' Use Find to locate occurrences of "PARENT_NM" in the first column
    Set foundCell = searchRange.Find(What:="PARENT_NM", LookIn:=xlValues, LookAt:=xlWhole)

    ' Collect all rows to delete in a range
    If Not foundCell Is Nothing Then
        firstAddress = foundCell.Address ' Store the first match to avoid infinite loop
        Do
            ' Add the found row to the delete range
            If deleteRows Is Nothing Then
                Set deleteRows = foundCell.EntireRow
            Else
                Set deleteRows = Union(deleteRows, foundCell.EntireRow)
            End If
            
            ' Look for the next occurrence
            Set foundCell = searchRange.Find(What:="PARENT_NM", After:=foundCell, LookIn:=xlValues, LookAt:=xlWhole)
        Loop While Not foundCell Is Nothing And foundCell.Address <> firstAddress
    End If

    ' Delete all collected rows at once (faster and avoids range shifting issues)
    If Not deleteRows Is Nothing Then deleteRows.Delete


' /*
' STEP 2: EXPORT DESIRED COLUMNS TO NEW WORKSHEET CALLED FilteredData
' */

    Dim colName As Variant
    Dim headers As Range
    Dim keepList As Variant
    Dim i As Long
    Dim workUnitCol As Range
    Dim lastRow As Long
    Dim targetCol As Long

    ' Create a new worksheet for the filtered data
    Set wsFilteredData = ThisWorkbook.Sheets.Add
    wsFilteredData.Name = "FilteredData" ' Change as needed

    ' Define the columns to keep
    keepList = Array("PARENT_NM", "CORP_CD", "WORK_UNIT_CD", "STMT_CNT", "INSERT_CNT", "REM_MC_CNT", "PLAN_TYPE_CD") ' Desired column names
    Set headers = wsSpecial.Rows(1) ' Assuming headers are in row 1

    ' Copy the keepList columns to the new worksheet
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

    ' Set reference for WORK_UNIT_CD column
    Set workUnitCol = wsFilteredData.Rows(1).Find("WORK_UNIT_CD")

    ' Validate that WORK_UNIT_CD column exists
    If workUnitCol Is Nothing Then
        MsgBox "Required column `WORK_UNIT_CD` not found!", vbExclamation
        Exit Sub
    End If

    ' Sort data by WORK_UNIT_CD
    lastRow = wsFilteredData.Cells(wsFilteredData.Rows.Count, workUnitCol.Column).End(xlUp).Row
    wsFilteredData.Sort.SortFields.Clear
    wsFilteredData.Sort.SortFields.Add key:=wsFilteredData.Columns(workUnitCol.Column), Order:=xlAscending ' Sort by WORK_UNIT_CD

    With wsFilteredData.Sort
        .SetRange wsFilteredData.UsedRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

' /*
' STEP 3: GET OUTERS BASED ON CORP_CD
' */

    Dim wsOutersKey As Worksheet
    Dim datasetCORPCol As Range, outerCol As Range
    Dim planTypeCol As Range
    Dim lastRowDataset As Long, lastRowOutersKey As Long
    Dim matchLength As Long
    Dim datasetCORPValue As String, planTypeValue As String, mappedOuter As String
    Dim sortedOuters(1 To 6) As Collection
    Dim entry As Variant

    ' Set OUTERSKEY worksheet
    On Error Resume Next
    Set wsOutersKey = ThisWorkbook.Sheets("outerskey") ' Update to your OUTERSKEY sheet name
    On Error GoTo 0
    
    ' If wsSpecial is missing, just exit
    If wsOutersKey Is Nothing Then
        MsgBox "`OUTERSKEY` worksheet is missing. Please add the `OUTERSKEY` worksheet.", vbInformation
        Exit Sub
    End If

    ' Find the relevant columns in the dataset
    Set datasetCORPCol = wsFilteredData.Rows(1).Find("CORP_CD")
    Set planTypeCol = wsFilteredData.Rows(1).Find("PLAN_TYPE_CD")
    Set outerCol = wsFilteredData.Rows(1).Find("OUTER")

    ' Check if OUTER column exists
    If outerCol Is Nothing Then
        Dim lastColumn As Long
        lastColumn = wsFilteredData.Cells(1, wsFilteredData.Columns.Count).End(xlToLeft).Column + 1
        wsFilteredData.Cells(1, lastColumn).Value = "OUTER"
        Set outerCol = wsFilteredData.Cells(1, lastColumn)
    End If

    ' Find the relevant columns in OUTERSKEY
    Dim outC5OuterCol As Range, outC4OuterCol As Range, outDLOuterCol As Range, outerCORPCol As Range
    Set outC5OuterCol = wsOutersKey.Rows(1).Find("C5_OUTER")
    Set outC4OuterCol = wsOutersKey.Rows(1).Find("C4_OUTER")
    Set outDLOuterCol = wsOutersKey.Rows(1).Find("DL_OUTER")
    Set outerCORPCol = wsOutersKey.Rows(1).Find("CORP_CD")

    ' Validate columns exist
    If datasetCORPCol Is Nothing Or planTypeCol Is Nothing Then
        MsgBox "One or more of the required columns `CORP_CD`, `PLAN_TYPE_CD` not found in special!", vbExclamation
        Exit Sub
    End If
    If outerCORPCol Is Nothing Or outC5OuterCol Is Nothing Or outC4OuterCol Is Nothing Or outDLOuterCol Is Nothing Then
        MsgBox "One or more of the required columns `CORP_CD`, `C5_OUTER`, `C4_OUTER`, `DL_OUTER` not found in OUTERSKEY!", vbExclamation
        Exit Sub
    End If

    ' Populate sortedOuters based on CORP_CD length
    lastRowOutersKey = wsOutersKey.Cells(wsOutersKey.Rows.Count, outerCORPCol.Column).End(xlUp).Row
    For matchLength = 1 To 6
        Set sortedOuters(matchLength) = New Collection
    Next matchLength
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
    lastRowDataset = wsFilteredData.Cells(wsFilteredData.Rows.Count, datasetCORPCol.Column).End(xlUp).Row
    For i = 2 To lastRowDataset
        datasetCORPValue = Trim(wsFilteredData.Cells(i, datasetCORPCol.Column).Value)
        planTypeValue = Trim(wsFilteredData.Cells(i, planTypeCol.Column).Value)
        mappedOuter = ""

        For matchLength = 6 To 1 Step -1
            For Each entry In sortedOuters(matchLength)
                If Left(datasetCORPValue, Len(entry(0))) = entry(0) Then
                    If planTypeValue = "V" Or planTypeValue = "F" Then
                        mappedOuter = entry(2) ' Use C4_OUTER
                        ' also highlight cells where PLAN_TYPE is V or F
                        wsFilteredData.Cells(i, planTypeCol.Column).Interior.Color = cPink
                        wsFilteredData.Cells(i, outerCol.Column).Interior.Color = cPink
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
        wsFilteredData.Cells(i, outerCol.Column).Value = mappedOuter
    Next i

    ' Apply formatting to wsFilteredData
    With wsFilteredData
        
        ' Left-align first column
        .Columns(8).HorizontalAlignment = xlLeft

        ' Shrink column widths where necessary
        Columns("A").ColumnWidth = 8
        Columns("B").ColumnWidth = 8
        Columns("C").ColumnWidth = 8
        Columns("D").ColumnWidth = 6.5
        Columns("E").ColumnWidth = 5
        Columns("F").ColumnWidth = 4
        Columns("G").ColumnWidth = 3
        Columns("H").ColumnWidth = 12

    ' Apply borders to all data cells
        With .Range(.Cells(1, 1), .Cells(lastRowDataset, 8)).Borders 'using 8 as we have 8 columns to border
            .LineStyle = xlContinuous
            .Color = vbBlack
            .Weight = xlThin
        End With
    End With
    
    ' /*
' STEP 4: HIGHLIGHT NEW ENTRIES
' */

    Dim wsPreviousFilteredData As Worksheet
    Dim latestWorkOrderCol As Range, previousWorkOrderCol As Range
    Dim arrLatestWorkOrders As Variant, arrPreviousWorkOrders As Variant
    Dim lastRowFiltered As Long, lastRowPrevious As Long
    Dim lastColFiltered As Long, lastColPrevious As Long
    Dim j As Long
    Dim isFound As Boolean
    Dim cell As Range

    ' Set the target worksheets
    On Error Resume Next
    Set wsPreviousFilteredData = ThisWorkbook.Sheets("previous")
    On Error GoTo 0
    
    ' If wsPreviousFilteredData is missing, just skip comparison
    If wsPreviousFilteredData Is Nothing Then
        MsgBox "`previous` worksheet is missing. Please rename `FilteredData` to `previous` and execute the formula again.", vbInformation
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
        MsgBox "`WORK_UNIT_CD` column not found!", vbExclamation
        Exit Sub
    End If

    ' Load WORK_UNIT_CD values into arrays
    arrLatestWorkOrders = wsFilteredData.Range(wsFilteredData.Cells(2, latestWorkOrderCol.Column), wsFilteredData.Cells(lastRowFiltered, latestWorkOrderCol.Column)).Value
    arrPreviousWorkOrders = wsPreviousFilteredData.Range(wsPreviousFilteredData.Cells(2, previousWorkOrderCol.Column), wsPreviousFilteredData.Cells(lastRowPrevious, previousWorkOrderCol.Column)).Value

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
            wsFilteredData.Cells(i + 1, datasetCORPCol.Column).Interior.Color = cBlue
        End If
    Next i

' /*
' STEP 5: HIGHLIGHT WORK ORDERS AND INSERTS WHERE INSERTS > 4
' */

    Dim insertCntCol As Range
    Dim insertCntValue As Variant
    
    ' Find the columns for WORK_UNIT_CD and INSERT_CNT
    Set workUnitCol = wsFilteredData.Rows(1).Find("WORK_UNIT_CD")
    Set insertCntCol = wsFilteredData.Rows(1).Find("INSERT_CNT")
    
    ' Validate the columns exist
    If workUnitCol Is Nothing Or insertCntCol Is Nothing Then
        MsgBox "One or more of the required columns `WORK_UNIT_CD`, `INSERT_CNT` not found!", vbExclamation
        Exit Sub
    End If
    
    ' Loop through each row to check the INSERT_CNT value
    For i = 2 To lastRow
        insertCntValue = wsFilteredData.Cells(i, insertCntCol.Column).Value
        
        ' Check if INSERT_CNT is greater than 9
        If IsNumeric(insertCntValue) And insertCntValue > 4 Then
            ' Highlight WORK_UNIT_CD 
            wsFilteredData.Cells(i, workUnitCol.Column).Interior.Color = cRed
            ' Highlight INSERT_CNT
            wsFilteredData.Cells(i, insertCntCol.Column).Interior.Color = cOrange
        End If
    Next i

' /*
' STEP 6: HIGHLIGHT REMAKES
' */

    Dim remCountCol As Range

    ' Find the REM_MC_CNT column
    Set remCountCol = wsFilteredData.Rows(1).Find("REM_MC_CNT")

    ' Validate that REM_MC_CNT column exists
    If remCountCol Is Nothing Then
        MsgBox "`REM_MC_CNT` column was not found!", vbExclamation
        Exit Sub
    End If

    ' Loop through each row in the REM_MC_CNT column
    For i = 2 To lastRow ' Assuming headers are in row 1
        If wsFilteredData.Cells(i, remCountCol.Column).Value <> "" Then
            ' Highlight the row up to the last column
            wsFilteredData.Range(wsFilteredData.Cells(i, 1), wsFilteredData.Cells(i, lastColumn)).Interior.Color = cYellow
        End If
    Next i

' /*
' STEP 7: CREATE A COLOUR KEY ON FilteredData
' */

    Dim startRow As Long, endRow As Long
    Dim keyDescriptions As Variant
    Dim keyColors As Variant
    Dim colorKeyRange As Range
    
    ' Find the first empty row below the data
    startRow = wsFilteredData.Cells(wsFilteredData.Rows.Count, 1).End(xlUp).Row + 4 ' 4 rows below the last row of data (taking into account the Heading Row)

    ' Calculate the heading row (directly above the color key)
    headingRow = startRow - 1
    
    ' Add the heading text
    With wsFilteredData.Range(wsFilteredData.Cells(headingRow, 1), wsFilteredData.Cells(headingRow, 3))
        .Merge ' Merge across columns A to C
        .Value = "Color Key" ' Set the heading text
        .HorizontalAlignment = xlCenter ' Center-align the text
        .VerticalAlignment = xlCenter
        .Font.Bold = True ' Make the text bold
        .Interior.Color = RGB(200, 200, 200) ' Optional: Light gray background color
    End With
    
    ' Define the descriptions and their corresponding colors
    keyDescriptions = Array("Remakes", _
                            "C4 Outers", _
                            "Work Orders with Inserts", _
                            "Inserts", _
                            "New Entries")
    keyColors = Array(cYellow, cPink, cRed, cOrange, cBlue)
    
    ' Add the color key
    For i = LBound(keyDescriptions) To UBound(keyDescriptions)
        ' Write the description
        wsFilteredData.Cells(startRow + i, 1).Value = keyDescriptions(i) ' Column A for descriptions
        
        ' Apply the background color to Column A
        wsFilteredData.Cells(startRow + i, 1).Interior.Color = keyColors(i)

        wsFilteredData.Range(wsFilteredData.Cells(startRow + i, 1), wsFilteredData.Cells(startRow + i, 3)).Merge
    Next i

    ' Find the first row of the color key
    endRow = wsFilteredData.Cells(wsFilteredData.Rows.Count, 1).End(xlUp).Row
    
    ' Combine the heading row and color key range
    Set colorKeyRange = wsFilteredData.Range(wsFilteredData.Cells(headingRow, 1), wsFilteredData.Cells(endRow, 3)) ' A to C, heading + color key
    
    ' Apply a thin black border to each row in the range
    For i = headingRow To endRow
        With wsFilteredData.Range(wsFilteredData.Cells(i, 1), wsFilteredData.Cells(i, 3)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0) ' Black border
        End With
    Next i

' /*
' STEP 8: CALCULATE A SUMMARY
' */

    Dim wsSummary As Worksheet
    Dim stmtCNCol As Range, remMCCol As Range
    Dim outerValue As String, stmtValue As Double, remMCValue As Variant
    Dim summaryRow As Long
    Dim idx As Long
    Dim outerArray() As Variant, stmtSumArray() As Double, stockArray() As Variant
    Dim stockLocation As String
    Dim foundOuter As Boolean
    
    ' Find the relevant columns
    Set stmtCNCol = wsFilteredData.Rows(1).Find("STMT_CNT")
    Set remMCCol = wsFilteredData.Rows(1).Find("REM_MC_CNT")
    ' Set outerCol = wsFilteredData.Rows(1).Find("OUTER") 'declared elsewhere in code
    ' Set planTypeCol = wsFilteredData.Rows(1).Find("PLAN_TYPE_CD") 'declared elsewhere in code
    
    ' Validate the columns exist in wsFilteredData
    If outerCol Is Nothing Or stmtCNCol Is Nothing Or remMCCol Is Nothing Or planTypeCol Is Nothing Then
        MsgBox "One or more of the required columns `OUTER`, `STMT_CNT`, `REM_MC_CNT`, `PLAN_TYPE_CD` not found!", vbExclamation
        Exit Sub
    End If
    
    ' Find the last rows
    lastRowDataset = wsFilteredData.Cells(wsFilteredData.Rows.Count, outerCol.Column).End(xlUp).Row
    lastRowOutersKey = wsOutersKey.Cells(wsOutersKey.Rows.Count, 1).End(xlUp).Row
    
    ' Initialize arrays for OUTER values, SUM values, and STOCK_LOCATION
    ReDim outerArray(1 To 1)
    ReDim stmtSumArray(1 To 1)
    ReDim stockArray(1 To 1)
    
    ' Loop through each row in wsFilteredData to calculate sums and map STOCK_LOCATION
    For i = 2 To lastRowDataset
        outerValue = wsFilteredData.Cells(i, outerCol.Column).Value
        stmtValue = wsFilteredData.Cells(i, stmtCNCol.Column).Value
        remMCValue = wsFilteredData.Cells(i, remMCCol.Column).Value
        planTypeValue = wsFilteredData.Cells(i, planTypeCol.Column).Value
        
        ' If REM_MC_CNT has a value, use it instead of STMT_CNT
        If Not IsEmpty(remMCValue) And IsNumeric(remMCValue) Then
            stmtValue = remMCValue
        End If

        If outerValue <> "" Then
            foundOuter = False
            stockLocation = ""
            
            ' Check if OUTER value already exists in the array
            For idx = 1 To UBound(outerArray)
                If outerArray(idx) = outerValue Then
                    stmtSumArray(idx) = stmtSumArray(idx) + stmtValue
                    foundOuter = True
                    Exit For
                End If
            Next idx

            ' If OUTER value not found, add new entry and determine STOCK_LOCATION
            If Not foundOuter Then
                ReDim Preserve outerArray(1 To UBound(outerArray) + 1)
                ReDim Preserve stmtSumArray(1 To UBound(stmtSumArray) + 1)
                ReDim Preserve stockArray(1 To UBound(stockArray) + 1)
                
                outerArray(UBound(outerArray)) = outerValue
                stmtSumArray(UBound(stmtSumArray)) = stmtValue
                
                ' Map STOCK_LOCATION based on OUTER and PLAN_TYPE_CD
                For idx = 2 To lastRowOutersKey
                    If planTypeValue = "V" Or planTypeValue = "F" Then
                        If wsOutersKey.Cells(idx, 3).Value = outerValue Then ' Match in C4_OUTER
                            stockLocation = wsOutersKey.Cells(idx, 6).Value ' C4_STOCK_LOCATION
                            Exit For
                        End If
                    Else
                        If wsOutersKey.Cells(idx, 2).Value = outerValue Then ' Match in C5_OUTER
                            stockLocation = wsOutersKey.Cells(idx, 5).Value ' C5_STOCK_LOCATION
                            Exit For
                        ElseIf wsOutersKey.Cells(idx, 4).Value = outerValue Then ' Match in DL_OUTER
                            stockLocation = wsOutersKey.Cells(idx, 7).Value ' DL_STOCK_LOCATION
                            Exit For
                        End If
                    End If
                Next idx
                
                stockArray(UBound(stockArray)) = stockLocation
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
    wsSummary.Cells(1, 3).Value = "STOCK_LOCATION"

    summaryRow = 2
    For idx = 1 To UBound(outerArray)
        wsSummary.Cells(summaryRow, 1).Value = outerArray(idx)
        wsSummary.Cells(summaryRow, 2).Value = stmtSumArray(idx)
        wsSummary.Cells(summaryRow, 3).Value = stockArray(idx)
        summaryRow = summaryRow + 1
    Next idx

    ' Apply formatting to wsSummary
    With wsSummary
        Dim summaryLastCol As Long, summaryLastRow As Long
        summaryLastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        summaryLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row

        ' Set headers bold and italic
        .Range(.Cells(1, 1), .Cells(1, summaryLastCol)).Font.Bold = True
        .Range(.Cells(1, 1), .Cells(1, summaryLastCol)).Font.Italic = True

        ' Adjust column widths
        .Columns.AutoFit

        ' Make Columns A and B slightly wider
        Columns("A").ColumnWidth = 15
        Columns("B").ColumnWidth = 10

        ' Left-align first and second columns
        .Columns(1).HorizontalAlignment = xlLeft
        .Columns(2).HorizontalAlignment = xlLeft

        ' Apply borders to all data cells
        With .Range(.Cells(1, 1), .Cells(summaryLastRow, summaryLastCol)).Borders
            .LineStyle = xlContinuous
            .Color = vbBlack
            .Weight = xlThin
        End With
    End With
    
    ' Sort data by OUTER
    wsSummary.Sort.SortFields.Clear
    wsSummary.Sort.SortFields.Add key:=wsSummary.Columns(1), Order:=xlAscending

    With wsSummary.Sort
        .SetRange wsSummary.UsedRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
    
End Sub