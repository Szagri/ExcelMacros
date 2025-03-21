Attribute VB_Name = "CopyMacroScript"
Sub CopyMacro(startColumn As Integer, endColumn As Integer, startRow As Integer, modifyColumns As Boolean, controlColumn As Integer)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim sourceRow As Long
    Dim currentValue As Variant, previousValue As Variant
    Dim firstDataFound As Boolean
    Dim clearedRows() As Boolean
    Dim colorRows As Boolean

    On Error GoTo ErrorHandler
    
    Set ws = ActiveWorkbook.Sheets(1)

    lastRow = ws.Cells(ws.Rows.Count, controlColumn).End(xlUp).Row

    ReDim clearedRows(1 To lastRow)
    
    colorRows = CopyMacroForm.chkColorRows.Value
    
    For j = endColumn To startColumn Step -1
        previousValue = ""
        firstDataFound = False
        sourceRow = 0

        For i = startRow To lastRow
            currentValue = ws.Cells(i, j).Value

            If Not IsEmpty(currentValue) Then
                If Not firstDataFound Then

                    firstDataFound = True
                    sourceRow = i
                    previousValue = currentValue
                ElseIf currentValue <> previousValue Then
                    If modifyColumns And Not clearedRows(i) Then
                        Call ClearRow(ws, i, j + 1, endColumn)
                        clearedRows(i) = True
                    End If
                    If colorRows Then
                        ws.Cells(i, j).Interior.Color = RGB(255, 255, 0)
                    End If
                    sourceRow = i
                    previousValue = currentValue
                End If
            ElseIf firstDataFound Then
                ws.Cells(i, j).Value = ws.Cells(sourceRow, j).Value
            End If
        Next i
    Next j

    MsgBox "Kopiowanie zakoñczone z uwzglêdnieniem parametrów u¿ytkownika!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Wyst¹pi³ b³¹d: " & Err.Description, vbExclamation
    Resume Next
End Sub

Sub ClearRow(ws As Worksheet, rowNum As Long, startCol As Integer, endCol As Integer)
    Dim col As Integer
    
    For col = startCol To endCol
        ws.Cells(rowNum, col).ClearContents
    Next col
End Sub
