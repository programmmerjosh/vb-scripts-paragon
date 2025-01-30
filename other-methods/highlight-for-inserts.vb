' This script does the following:
'   1. Highlights every WORK_UNIT_CD and INSERT_CNT (2 different colours) where the INSERT_CNT is greater than 9.

'IMPORTANT NOTE: Don't forget to change the sheet name to, "Special1" before running this script

Sub HighlightForInserts()
    Dim ws As Worksheet
    Dim workUnitCol As Range, insertCntCol As Range
    Dim lastRow As Long
    Dim i As Long
    Dim insertCntValue As Variant
    
    ' Set the dataset worksheet
    Set ws = ThisWorkbook.Sheets("Special1") ' Update to your sheet name
    
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
    
    ' MsgBox "Cells highlighted successfully based on INSERT_CNT criteria.", vbInformation
End Sub
