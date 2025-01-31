Sub MergeMySheets()
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim lastRow As Long, targetRow As Long
    Dim wb As Workbook, ColCnt As Long
    Dim sheetNames As Variant, srcRng As Range
    Dim i As Integer
    
    ' Define the sheet names to copy from
    sheetNames = Array("s1", "s2", "s3", "s4", "s5", "s6", "s7", "s8") ' Update as needed
    
    ' Create a new worksheet for the filtered data
    Set wb = ThisWorkbook
    
    ' Check if the sheet exists before proceeding
    On Error Resume Next
    Set wsTarget = wb.Sheets("special")
    On Error GoTo 0 ' Reset error handling
    If Not wsTarget Is Nothing Then
        ' Clear existing data in target sheet
        wsTarget.Cells.Clear
    Else
        Set wsTarget = wb.Sheets.Add
        wsTarget.Name = "special" ' Change as needed
    End If
    
    targetRow = 1 ' Start pasting from the first row
    Dim firstSheet As Boolean: firstSheet = True ' Flag to track the first sheet
    
    ' Loop through the defined sheet names
    For i = LBound(sheetNames) To UBound(sheetNames)
        ' Check if the sheet exists before proceeding
        On Error Resume Next
        Set wsSource = wb.Sheets(sheetNames(i))
        On Error GoTo 0 ' Reset error handling
        
        ' If the sheet does not exist, skip it
        If Not wsSource Is Nothing Then
            ' Find last used row in source sheet
            lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
            ColCnt = wsSource.UsedRange.Columns.Count
            ' Only copy if the sheet contains data
            If lastRow > 0 Then
                If firstSheet Then
                    ' First sheet: Copy everything (including headers)
                    Set srcRng = wsSource.Range("A1", wsSource.Cells(lastRow, ColCnt))
                    firstSheet = False ' Mark that we've processed the first sheet
                Else
                    ' Other sheets: Exclude the header (start from row 2)
                    If lastRow > 1 Then
                        Set srcRng = wsSource.Range("A2", wsSource.Cells(lastRow, ColCnt))
                    Else
                        ' If only header exists, skip this sheet
                    End If
                End If
                If Not srcRng Is Nothing Then
                    ' Paste values into target sheet
                    With srcRng
                        wsTarget.Cells(targetRow, 1).Resize(.Rows.Count, .Columns.Count).Value = .Value
                    End With
                End If
                ' Update next target row
                targetRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row + 1
            End If
        End If
        
        ' Reset wsSource for the next loop
        Set wsSource = Nothing
        Set srcRng = Nothing
    Next i
    
    ' MsgBox "Data merged successfully! Now running the primary script.", vbInformation
    
    ' Call the primary script on the merged data
    Call FilterDataAndCreateSummary
    
End Sub