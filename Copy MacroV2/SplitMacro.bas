Attribute VB_Name = "SplitMacro"
Sub SplitData()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentValue As String

    Set ws = ActiveWorkbook.Sheets(1)
    
    lastRow = ws.Cells(ws.Rows.Count, 16).End(xlUp).Row

    For i = 4 To lastRow

        currentValue = CStr(ws.Cells(i, 1).Value)
        If Len(currentValue) >= 4 Then
            ws.Cells(i, 1).Value = "'" & Mid(currentValue, 1, 2)
            ws.Cells(i, 3).Value = "'" & Mid(currentValue, 3, 2)
        End If

        currentValue = CStr(ws.Cells(i, 4).Value)
        If Len(currentValue) >= 6 Then
            ws.Cells(i, 6).Value = "'" & Mid(currentValue, 5, 2)
        End If

        currentValue = CStr(ws.Cells(i, 7).Value)
        If Len(currentValue) >= 7 Then
            ws.Cells(i, 9).Value = "'" & Mid(currentValue, 7, 1)
        End If

        currentValue = CStr(ws.Cells(i, 10).Value)
        If Len(currentValue) >= 9 Then
            ws.Cells(i, 12).Value = "'" & Mid(currentValue, 8, 2)
        End If

        currentValue = CStr(ws.Cells(i, 13).Value)
        If Len(currentValue) >= 14 Then
            ws.Cells(i, 15).Value = "'" & Mid(currentValue, 10, 2)
        End If

    Next i

    ws.Columns(13).Delete
    ws.Columns(10).Delete
    ws.Columns(7).Delete
    ws.Columns(4).Delete
    ws.Columns(1).Delete

    ws.Columns(11).Insert Shift:=xlToRight
    ws.Columns(11).Insert Shift:=xlToRight
    
    DoEvents

    For i = 4 To lastRow
        currentValue = CStr(ws.Cells(i, 13).Value)
        If Len(currentValue) >= 14 Then
            ws.Cells(i, 11).Value = "'" & Mid(currentValue, 12, 2)
            ws.Cells(i, 12).Value = "'" & Mid(currentValue, 14, 1)
        End If
    Next i

    MsgBox "Rozdzielanie danych zakoñczone, kolumny dodane i uzupe³nione!", vbInformation

End Sub
