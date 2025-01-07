' This script does the following:
'   1. Gets the outer (based on the CORP_CD) for each work-order on the active sheet.

'IMPORTANT NOTE: Don't forget to change the sheet name to, "Special1" before running this script.

Sub GetOuters()
    Dim wsDataset As Worksheet
    Dim wsOutersKey As Worksheet
    Dim datasetCORPCol As Range, outerCol As Range
    Dim outC5OuterCol As Range, outC4OuterCol As Range, outDLOuterCol As Range
    Dim planTypeCol As Range
    Dim lastRowDataset As Long, lastRowOutersKey As Long
    Dim i As Long, matchLength As Long
    Dim datasetCORPValue As String, planTypeValue As String, mappedOuter As String
    Dim sortedOuters(1 To 6) As Collection ' Array of collections to group by length
    Dim entry As Variant

    ' Set the dataset and OUTERSKEY worksheets
    Set wsDataset = ThisWorkbook.Sheets("Special1") ' Update to your dataset sheet name
    Set wsOutersKey = ThisWorkbook.Sheets("OUTERSKEY") ' Update to your OUTERSKEY sheet name

    ' Find the relevant columns in the dataset
    Set datasetCORPCol = wsDataset.Rows(1).Find("CORP_CD")
    Set planTypeCol = wsDataset.Rows(1).Find("PLAN_TYPE")

    ' Check if OUTER column exists
    Set outerCol = wsDataset.Rows(1).Find("OUTER")

    ' If OUTER column doesn't exist, add it at the next empty column
    If outerCol Is Nothing Then
        Dim lastColumn As Long
        lastColumn = wsDataset.Cells(1, wsDataset.Columns.Count).End(xlToLeft).Column + 1
        wsDataset.Cells(1, lastColumn).Value = "OUTER"
        Set outerCol = wsDataset.Cells(1, lastColumn) ' Now set outerCol to the newly created column
    End If

    ' Find the relevant columns in OUTERSKEY
    Set outC5OuterCol = wsOutersKey.Rows(1).Find("C5_OUTER")
    Set outC4OuterCol = wsOutersKey.Rows(1).Find("C4_OUTER")
    Set outDLOuterCol = wsOutersKey.Rows(1).Find("DL_OUTER")
    Set outerCORPCol = wsOutersKey.Rows(1).Find("CORP_CD")

    ' Validate columns exist
    If datasetCORPCol Is Nothing Or planTypeCol Is Nothing Then
        MsgBox "Required columns (CORP_CD, PLAN_TYPE) not found in the dataset!", vbExclamation
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
    lastRowDataset = wsDataset.Cells(wsDataset.Rows.Count, datasetCORPCol.Column).End(xlUp).Row
    For i = 2 To lastRowDataset
        datasetCORPValue = Trim(wsDataset.Cells(i, datasetCORPCol.Column).Value)
        planTypeValue = Trim(wsDataset.Cells(i, planTypeCol.Column).Value)
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
        wsDataset.Cells(i, outerCol.Column).Value = mappedOuter
    Next i

    ' Convert OUTER column to text format
    For i = 2 To lastRowDataset
        wsDataset.Cells(i, outerCol.Column).Value = CStr(wsDataset.Cells(i, outerCol.Column).Value) ' Convert to string
    Next i

    ' Apply Text format to the entire OUTER column
    wsDataset.Columns(outerCol.Column).NumberFormat = "@"

    MsgBox "OUTERSKEY mapping to DATASET completed!", vbInformation
End Sub
