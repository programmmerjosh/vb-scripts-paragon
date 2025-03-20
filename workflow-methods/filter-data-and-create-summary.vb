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

' /*
' STEP 1: EXPORT DESIRED COLUMNS TO NEW WORKSHEET CALLED FilteredData
' */

    Dim colName As Variant
    Dim headers As Range
    Dim keepList As Variant
    Dim i As Long
    Dim workUnitCol As Range
    Dim lastRow As Long
    Dim targetCol As Long

    ' Check if the sheet exists before proceeding
    On Error Resume Next
    Set wsFilteredData = ThisWorkbook.Sheets("FilteredData")
    On Error GoTo 0 ' Reset error handling
    If Not wsFilteredData Is Nothing Then
        'Warn user that FilteredData already exists!
        MsgBox "`FilteredData` already exists!! For this script to execute properly, either delete `FilteredData` or rename it to `previous`", vbExclamation
        Exit Sub
    Else
        ' Create a new worksheet for the filtered data
        Set wsFilteredData = ThisWorkbook.Sheets.Add
        wsFilteredData.Name = "FilteredData" ' Change as needed
    End If

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
' STEP 2: GET OUTERS BASED ON CORP_CD
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
                        ' /*
                        ' Side STEP: HIGHLIGHT C4 OUTERS
                        ' */

                        ' also highlight cells where PLAN_TYPE is V or F (and OUTER in the same row)
                        ' wsFilteredData.Cells(i, planTypeCol.Column).Interior.Color = cPink
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

        ' /*
        ' Side STEP: HIGHLIGHT OUTERS WE ALWAYS NEED TO ORDER (even when they have zero inserts)
        ' */

        Dim outersToOrder As Variant
        Dim myOuter As String
        Dim matchFound As Boolean
        Dim count As Integer

        ' Define the array of string values
        outersToOrder = Array("50023", "BCY03", "BCORPC5AIR", "BARCLPC52", "GCRP5254TNT", "EOP39TNT", "BSMTNT")

        ' Define the variable to compare
        myOuter = wsFilteredData.Cells(i, outerCol.Column).Value
        matchFound = False

        ' Loop through the array and compare each value to myOuter
        For count = LBound(outersToOrder) To UBound(outersToOrder)
            If StrComp(outersToOrder(count), myOuter, vbTextCompare) = 0 Then
                ' wsFilteredData.Cells(i, outerCol.Column).Interior.Color = cOrange
                wsFilteredData.Cells(i, workUnitCol.Column).Interior.Color = cOrange
                matchFound = True
                Exit For ' Exit the loop since we found a match
            End If
        Next count
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
' STEP 3: HIGHLIGHT WORK ORDERS AND INSERTS WHERE INSERTS > 4
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
            ' wsFilteredData.Cells(i, insertCntCol.Column).Interior.Color = cRed
        End If
    Next i

' /*
' STEP 4: HIGHLIGHT REMAKES
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
' STEP 5: CREATE A COLOUR KEY ON FilteredData
' */

    Dim startRow As Long, endRow As Long
    Dim keyDescriptions As Variant
    Dim keyColors As Variant
    Dim colorKeyRange As Range
    Dim numberOfColsToMerge As Long

    numberOfColsToMerge = 6
    
    ' Find the first empty row below the data
    startRow = wsFilteredData.Cells(wsFilteredData.Rows.Count, 1).End(xlUp).Row + 4 ' 4 rows below the last row of data (taking into account the Heading Row)

    ' Calculate the heading row (directly above the color key)
    headingRow = startRow - 1
    
    ' Add the heading text
    With wsFilteredData.Range(wsFilteredData.Cells(headingRow, 1), wsFilteredData.Cells(headingRow, numberOfColsToMerge))
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
                            "Work Orders Of Jobs With Inserts", _
                            "Work Orders Of Jobs With Outers We Should Order (0 Inserts)", _
                            "New Entries")
    keyColors = Array(cYellow, cPink, cRed, cOrange, cBlue)
    
    ' Add the color key
    For i = LBound(keyDescriptions) To UBound(keyDescriptions)
        ' Write the description
        wsFilteredData.Cells(startRow + i, 1).Value = keyDescriptions(i) ' Column A for descriptions
        
        ' Apply the background color to Column A
        wsFilteredData.Cells(startRow + i, 1).Interior.Color = keyColors(i)

        wsFilteredData.Range(wsFilteredData.Cells(startRow + i, 1), wsFilteredData.Cells(startRow + i, numberOfColsToMerge)).Merge
    Next i

    ' Find the first row of the color key
    endRow = wsFilteredData.Cells(wsFilteredData.Rows.Count, 1).End(xlUp).Row
    
    ' Combine the heading row and color key range
    Set colorKeyRange = wsFilteredData.Range(wsFilteredData.Cells(headingRow, 1), wsFilteredData.Cells(endRow, numberOfColsToMerge)) 
    
    ' Apply a thin black border to each row in the range
    For i = headingRow To endRow
        With wsFilteredData.Range(wsFilteredData.Cells(i, 1), wsFilteredData.Cells(i, numberOfColsToMerge)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0) ' Black border
        End With
    Next i

' /*
' STEP 6: CALCULATE A SUMMARY
' */

    Dim stmtCNCol As Range, remMCCol As Range
    Dim summaryStartRow As Long
    Dim outerValue As String, stmtValue As Double, remMCValue As Variant
    Dim summaryData As Collection, key As Variant
    Dim idx As Long
    Dim outerArray() As Variant, stmtSumArray() As Double, stockArray() As Variant
    Dim stockLocation As String
    Dim foundOuter As Boolean
    Dim summaryEndRow As Long
    Dim mergedRange As Range
    
    ' Find the relevant columns
    ' Set outerCol = wsFilteredData.Rows(1).Find("OUTER")
    Set stmtCNCol = wsFilteredData.Rows(1).Find("STMT_CNT")
    Set remMCCol = wsFilteredData.Rows(1).Find("REM_MC_CNT")
    ' Set planTypeCol = wsFilteredData.Rows(1).Find("PLAN_TYPE_CD")
    
    ' Validate the columns exist in wsFilteredData
    If outerCol Is Nothing Or stmtCNCol Is Nothing Or remMCCol Is Nothing Or planTypeCol Is Nothing Then
        MsgBox "Required columns (OUTER, STMT_CNT, REM_MC_CNT, PLAN_TYPE_CD) not found!", vbExclamation
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

    ' Determine the start of the summary
    summaryStartRow = lastRowDataset + 11

    ' Write the summary data
    wsFilteredData.Cells(summaryStartRow, 1).Value = "OUTER"
    wsFilteredData.Cells(summaryStartRow, 3).Value = "SUM"
    wsFilteredData.Cells(summaryStartRow, 4).Value = "STOCK_LOCATION"

    For idx = 1 To UBound(outerArray)
        rowIdx = summaryStartRow + idx

        wsFilteredData.Cells(rowIdx, 1).Value = outerArray(idx)

        ' Write SUM
        wsFilteredData.Cells(rowIdx, 3).Value = stmtSumArray(idx)

        wsFilteredData.Cells(rowIdx, 4).Value = stockArray(idx)

        ' Delete rows where SUM is zero
        If stmtSumArray(idx) = 0 Then ' Check if SUM (Column B) is 0
            wsFilteredData.Rows(rowIdx).Delete
        End If
    Next idx

    ' Determine the end of the summary
    summaryEndRow = summaryStartRow + UBound(outerArray)

    ' Sort the summary range
    With wsFilteredData.Sort
        .SortFields.Clear  ' Clear any previous sort fields
        .SortFields.Add key:=wsFilteredData.Columns(1), Order:=xlAscending ' Sort by OUTER column (Column A)
        
        .SetRange wsFilteredData.Range(wsFilteredData.Cells(summaryStartRow, 1), wsFilteredData.Cells(summaryEndRow, 4))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    ' Apply styling to the summary
    summaryLastCol = 7 ' Column G
    With wsFilteredData
        ' Set headers bold and italic
        .Range(.Cells(summaryStartRow, 1), .Cells(summaryStartRow, summaryLastCol)).Font.Bold = True
        .Range(.Cells(summaryStartRow, 1), .Cells(summaryStartRow, summaryLastCol)).Font.Italic = True

        .Range(.Cells(summaryStartRow, 1), .Cells(summaryEndRow, 1)).HorizontalAlignment = xlLeft

        ' Apply borders to the summary range
        With .Range(.Cells(summaryStartRow, 1), .Cells(summaryEndRow, summaryLastCol)).Borders
            .LineStyle = xlContinuous
            .Color = vbBlack
            .Weight = xlThin
        End With
    End With

    ' Merge summary heading columns
    wsFilteredData.Range(wsFilteredData.Cells(summaryStartRow, 1), wsFilteredData.Cells(summaryStartRow, 2)).Merge
    wsFilteredData.Range(wsFilteredData.Cells(summaryStartRow, 4), wsFilteredData.Cells(summaryStartRow, 7)).Merge

     For idx = 1 To UBound(outerArray)
        rowIdx = summaryStartRow + idx

        ' Merge columns for OUTER
        wsFilteredData.Range(wsFilteredData.Cells(rowIdx, 1), wsFilteredData.Cells(rowIdx, 2)).Merge

        ' Merge and write STOCK_LOCATION
        wsFilteredData.Range(wsFilteredData.Cells(rowIdx, 4), wsFilteredData.Cells(rowIdx, 7)).Merge
    Next idx

' /*
' STEP 7: DELETE `special` WORKSHEET AS WE WILL NO LONGER BE NEEDING IT.
' */
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("special").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

' /*
' STEP 8: HIGHLIGHT NEW ENTRIES (which will only execute if the 'previous' worksheet exists) 
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
        MsgBox "The script has run successfully!!", vbInformation
        MsgBox "Important Note: `previous` worksheet is missing. Rename `FilteredData` to `previous` before you run this script again to see the new entries.", vbInformation
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
' STEP 9: Proposed new step: Find a list of "enclosed" work orders
        ' NOTE: these will not be definate because if a work order is missing from our new-list, it COULD mean 
        ' that the work order has been pushed to the following day's list to be enclosed 
        ' unless we have already included tomorrow's list in our calculation, then it's almost certainly been enclosed.
' */

' /*
' STEP 10: DELETE `previous` WORKSHEET AS WE WILL NO LONGER BE NEEDING IT.
' */
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("previous").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    MsgBox "The script has run successfully!!", vbInformation
    
End Sub