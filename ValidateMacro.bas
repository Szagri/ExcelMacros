Attribute VB_Name = "ValidateMacro"
Sub ValidateEmails()
    Dim ws As Worksheet
    Dim emailCols As Variant
    Dim colIndex As Integer
    Dim lastRow As Long, i As Long
    Dim cell As Range
    Dim lastCol As Integer
    Dim foundCols As Object
    Dim colName As Variant
    
    Set ws = ActiveSheet
    emailCols = Array("E_MAIL_BJS", "E_MAIL_OZS_BJS", "E_MAIL_SPR_BJS", "OSSMAIL_PS_BJS", "E_MAIL_OSS_UR", "E_MAIL_JEDN_UR")
    
    lastCol = ws.UsedRange.Columns.Count
    
    Set foundCols = CreateObject("Scripting.Dictionary")
    
    For colIndex = 1 To lastCol
        For Each colName In emailCols
            If ws.Cells(1, colIndex).Value = colName Then
                If Not foundCols.Exists(colName) Then
                    foundCols.Add colName, colIndex
                End If
                Exit For
            End If
        Next colName
    Next colIndex
    
    For Each colName In foundCols.Keys
        colIndex = foundCols(colName)
        lastRow = ws.Cells(Rows.Count, colIndex).End(xlUp).Row
        
        For i = 2 To lastRow
            Set cell = ws.Cells(i, colIndex)
            
            cell.Value = CleanEmail(cell.Value)
             
            If cell.Value = "brak@brak.pl" Then
                cell.Value = ""
            ElseIf Not IsValidEmail(cell.Value) Then
                cell.Interior.Color = RGB(255, 182, 193)
            Else
                cell.Interior.ColorIndex = xlNone
            End If
        Next i
    Next colName
    
    MsgBox "Weryfikacja e-maili zakoñczona!", vbInformation
End Sub

Function CleanEmail(email As String) As String
    Dim separators As Variant
    Dim i As Integer
    Dim emailParts As Variant
    Dim cleanedEmail As String
    
    email = Trim(email)
    email = Replace(email, Chr(160), "")
    email = Replace(email, ChrW(&HAD), "")
    email = Replace(email, ChrW(9), "")
    email = Replace(email, ChrW(8203), "")
    email = Replace(email, vbCr, "")
    email = Replace(email, vbLf, "")
    email = Replace(email, Chr(58), "")
    email = Replace(email, Chr(34), "")
    
    separators = Array("; ", ";", " ,", ",", vbCr, vbLf)
    
    For i = LBound(separators) To UBound(separators)
        email = Replace(email, separators(i), ";")
    Next i
    
    emailParts = Split(email, ";")
    cleanedEmail = ""
    
    For i = LBound(emailParts) To UBound(emailParts)
        If Trim(emailParts(i)) <> "" Then
            If cleanedEmail <> "" Then cleanedEmail = cleanedEmail & ";"
            cleanedEmail = cleanedEmail & Trim(emailParts(i))
        End If
    Next i
    
    CleanEmail = cleanedEmail
End Function

Function IsValidEmail(email As String) As Boolean
    Dim regExp As Object
    Dim emailParts As Variant
    Dim i As Integer
    Dim singleEmail As String
    Dim atPos As Integer
    Dim hasError As Boolean
    Dim forbiddenWords As Variant
    Dim word As Variant
    
    If email = "" Then
        IsValidEmail = True
        Exit Function
    End If
    
    forbiddenWords = Array("b³¹d", "blad", "b³ad", "bl¹d", "brak", "Brak")
    
    emailParts = Split(email, ";")
    
    Set regExp = CreateObject("VBScript.RegExp")
    With regExp
        .Pattern = "^[A-Za-z0-9¹æê³ñóœŸ¿¥ÆÊ£ÑÓŒ¯._-]+@[A-Za-z0-9¹æê³ñóœŸ¿¥ÆÊ£ÑÓŒ¯._-]+\.[A-Za-z]{2,}$"
        .IgnoreCase = True
        .Global = False
    End With
    
    hasError = False
    
    For i = LBound(emailParts) To UBound(emailParts)
        singleEmail = Trim(emailParts(i))
        
        If singleEmail = "" Then GoTo NextEmail
        
        atPos = InStr(1, singleEmail, "@")
        
        If atPos = 0 Or InStr(atPos + 1, singleEmail, "@") > 0 Then
            hasError = True
            Exit For
        End If
        
        For Each word In forbiddenWords
            If InStr(1, LCase(singleEmail), LCase(word), vbTextCompare) > 0 Then
                hasError = True
                Exit For
            End If
        Next word
        
        If Not regExp.Test(singleEmail) Then
            hasError = True
            Exit For
        End If

NextEmail:
    Next i
    
    IsValidEmail = Not hasError
    
    Set regExp = Nothing
End Function
