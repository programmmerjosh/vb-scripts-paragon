' This script does the following:
'   1. Highlights every row (yellow) that has a value in the REM_MC_CNT field.

'IMPORTANT NOTE: Don't forget to change the sheet name to, "Special1" before running this script

Sub HighlightRemakes()
    Dim ws As Worksheet
    Dim remCountCol As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long

    ' Set worksheet
    Set ws = ThisWorkbook.Sheets("Special1") ' Update to your sheet name

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

    ' MsgBox "Highlighting complete!", vbInformation
End Sub

