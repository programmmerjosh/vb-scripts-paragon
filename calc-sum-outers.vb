Sub CalcSumOutersWithStockLocation()
    Dim wsDataset As Worksheet, wsOutersKey As Worksheet, wsSummary As Worksheet
    Dim outerCol As Range, stmtCNCol As Range, remMCCol As Range, planTypeCol As Range
    Dim stockLocationCol As Range
    Dim lastRowDataset As Long, lastRowOutersKey As Long
    Dim outerValue As String, stmtValue As Double, remMCValue As Variant
    Dim planTypeValue As String
    Dim summaryData As Collection, key As Variant
    Dim summaryRow As Long
    Dim i As Long, idx As Long
    Dim outerArray() As Variant, stmtSumArray() As Double, stockArray() As Variant
    Dim stockLocation As String
    Dim foundOuter As Boolean

    ' Set the worksheets
    Set wsDataset = ThisWorkbook.Sheets("FilteredData") ' Update to your dataset sheet name
    Set wsOutersKey = ThisWorkbook.Sheets("OUTERSKEY") ' Update to your OutersKey sheet name
    
    ' Find the relevant columns
    Set outerCol = wsDataset.Rows(1).Find("OUTER")
    Set stmtCNCol = wsDataset.Rows(1).Find("STMT_CNT")
    Set remMCCol = wsDataset.Rows(1).Find("REM_MC_CNT")
    Set planTypeCol = wsDataset.Rows(1).Find("PLAN_TYPE_CD")
    
    ' Validate the columns exist in wsDataset
    If outerCol Is Nothing Or stmtCNCol Is Nothing Or remMCCol Is Nothing Or planTypeCol Is Nothing Then
        MsgBox "Required columns (OUTER, STMT_CNT, REM_MC_CNT, PLAN_TYPE_CD) not found!", vbExclamation
        Exit Sub
    End If
    
    ' Find the last rows
    lastRowDataset = wsDataset.Cells(wsDataset.Rows.Count, outerCol.Column).End(xlUp).Row
    lastRowOutersKey = wsOutersKey.Cells(wsOutersKey.Rows.Count, 1).End(xlUp).Row
    
    ' Initialize arrays for OUTER values, SUM values, and STOCK_LOCATION
    ReDim outerArray(1 To 1)
    ReDim stmtSumArray(1 To 1)
    ReDim stockArray(1 To 1)
    
    ' Loop through each row in wsDataset to calculate sums and map STOCK_LOCATION
    For i = 2 To lastRowDataset
        outerValue = wsDataset.Cells(i, outerCol.Column).Value
        stmtValue = wsDataset.Cells(i, stmtCNCol.Column).Value
        remMCValue = wsDataset.Cells(i, remMCCol.Column).Value
        planTypeValue = wsDataset.Cells(i, planTypeCol.Column).Value
        
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

    MsgBox "Summary with STOCK_LOCATION created successfully in the 'Summary' worksheet.", vbInformation
End Sub
