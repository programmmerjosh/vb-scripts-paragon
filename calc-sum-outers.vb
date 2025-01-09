' This script does the following:
'   1. Creates a Summary worksheet calculating the total number of outers needed for each job on the active sheet.

'IMPORTANT NOTE: Don't forget to change the sheet name to, "Special1" before running this script. 
' And don't forget to run the GetOuters() script before this one.

Sub CalcSumOuters()
    Dim ws As Worksheet
    Dim wsSummary As Worksheet
    Dim outerCol As Range
    Dim stmtCNCol As Range, remMCCol As Range
    Dim lastRow As Long
    Dim outerValue As String
    Dim stmtValue As Double, remMCValue As Variant
    Dim i As Long
    Dim summaryData As Collection
    Dim key As Variant
    Dim summaryRow As Long
    Dim outerArray() As Variant
    Dim stmtSumArray() As Double
    Dim foundOuter As Boolean
    Dim idx As Long
    
    ' Set the dataset worksheet
    Set ws = ThisWorkbook.Sheets("Special1") ' Update to your dataset sheet name
    
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

    ' MsgBox "Summary of STMT_CNT by OUTER created successfully in the 'Summary' worksheet.", vbInformation
End Sub
