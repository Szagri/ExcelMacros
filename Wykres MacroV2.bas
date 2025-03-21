Attribute VB_Name = "ChartGen"
Sub ChartGenerator()
    Dim ws As Worksheet
    Dim wykresWs As Worksheet
    Dim daneStart As Range
    Dim cell As Range
    Dim i As Integer
    Dim wykres As Chart
    Dim lastRow As Long
    Dim outputRow As Long
    Dim rowIndex As Integer
    Dim wykresNr As Integer
    Dim odpowiedzi As Object
    Dim odpowiedziSlownik As Object
    Dim kolor As Long
    Dim odp As Variant
    Dim series As series
    Dim pointIndex As Integer
    Dim sortedOdpowiedzi As Object
    Dim key As Variant
    
    ' Aktywny arkusz
    Set ws = ActiveWorkbook.ActiveSheet
    
    ' Usuwanie akrusza wykres
    Application.DisplayAlerts = False
    On Error Resume Next
    Sheets("Wykresy").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' Tworzenie akrusza wykresy
    Set wykresWs = Worksheets.Add
    wykresWs.Name = "Wykresy"
    
    ' S�ownik odpowiedzi oraz dodawanie do niego warto�ci
    Set odpowiedziSlownik = CreateObject("Scripting.Dictionary")
    
    odpowiedziSlownik.Add "Zupe�nie si� nie zgadzam", Array(RGB(234, 67, 53), 5)
    odpowiedziSlownik.Add "Bardzo niski", Array(RGB(234, 67, 53), 5)
    odpowiedziSlownik.Add "Nie znam tego przepisu", Array(RGB(234, 67, 53), 5)
    odpowiedziSlownik.Add "Nigdy", Array(RGB(234, 67, 53), 5)
    
    odpowiedziSlownik.Add "Nie zgadzam si�", Array(RGB(255, 109, 1), 4)
    odpowiedziSlownik.Add "Rzadko", Array(RGB(255, 109, 1), 4)
    odpowiedziSlownik.Add "Niski", Array(RGB(255, 109, 1), 4)
    
    odpowiedziSlownik.Add "Nie mam zdania", Array(RGB(251, 188, 4), 3)
    odpowiedziSlownik.Add "Wiem, �e taki przepis istnieje", Array(RGB(251, 188, 4), 3)
    odpowiedziSlownik.Add "Niezbyt wysoki", Array(RGB(251, 188, 4), 3)
    odpowiedziSlownik.Add "Czasami", Array(RGB(251, 188, 4), 3)
    
    odpowiedziSlownik.Add "Cz�sto", Array(RGB(66, 133, 244), 2)
    odpowiedziSlownik.Add "Wysoki", Array(RGB(66, 133, 244), 2)
    odpowiedziSlownik.Add "Raczej si� zgadzam", Array(RGB(66, 133, 244), 2)
    
    odpowiedziSlownik.Add "Ca�kowicie si� zgadzam", Array(RGB(52, 168, 83), 1)
    odpowiedziSlownik.Add "Bardzo cz�sto", Array(RGB(52, 168, 83), 1)
    odpowiedziSlownik.Add "Bardzo wysoki", Array(RGB(52, 168, 83), 1)
    odpowiedziSlownik.Add "Ten przepis jest przestrzegany i funkcjonuje w�a�ciwie", Array(RGB(52, 168, 83), 1)
    
    outputRow = 1
    wykresNr = 1
    
    For i = 2 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        Set daneStart = ws.Range(ws.Cells(2, i), ws.Cells(ws.Rows.Count, i).End(xlUp))
        Set odpowiedzi = CreateObject("Scripting.Dictionary")
        
        ' Pozyskiwanie z tytu�u pytania miedzy "[]"
        rawTitle = ws.Cells(1, i).Value
        startPos = InStr(rawTitle, "[")
        endPos = InStr(rawTitle, "]")
        cleanedTitle = Mid(rawTitle, startPos + 1, endPos - startPos - 1)
        
        ' Zliczamy odpowiedzi w danych
        For Each cell In daneStart
            If Not odpowiedzi.exists(cell.Value) And cell.Value <> "" Then
                odpowiedzi.Add cell.Value, 1
            ElseIf cell.Value <> "" Then
                odpowiedzi(cell.Value) = odpowiedzi(cell.Value) + 1
            End If
        Next cell
        
        ' Sortowanie legendy wed�ug klucza w s�owniku
        Set sortedOdpowiedzi = CreateObject("Scripting.Dictionary")
        For Each key In odpowiedziSlownik.keys
            If odpowiedzi.exists(key) Then
                sortedOdpowiedzi.Add key, odpowiedzi(key)
            End If
        Next key
        
        ' Tworzenie tabeli
        wykresWs.Cells(outputRow, 1).Value = ws.Cells(1, i).Value
        wykresWs.Cells(outputRow, 1).Font.Bold = True
        outputRow = outputRow + 1
        rowIndex = outputRow
        
        ' Wypelnianie tabeli danymi na podstawie kt�rych b�dzie tworzony wykres
        For Each odp In sortedOdpowiedzi.keys
            wykresWs.Cells(rowIndex, 1).Value = odp
            wykresWs.Cells(rowIndex, 2).Value = sortedOdpowiedzi(odp)
            rowIndex = rowIndex + 1
        Next odp
        
        ' Tworzenie wykresu
        Set wykres = wykresWs.Shapes.AddChart2(251, xlPie).Chart
        wykres.SetSourceData wykresWs.Range(wykresWs.Cells(outputRow, 1), wykresWs.Cells(rowIndex - 1, 2))
        
        ' Tytu�
        wykres.ChartTitle.Text = wykresNr & ". " & cleanedTitle
        wykresNr = wykresNr + 1
        
        ' Rozmieszczenie
        wykres.Parent.Left = wykresWs.Cells(outputRow, 4).Left
        wykres.Parent.Top = wykresWs.Cells(rowIndex + 2, 1).Top
        wykres.Parent.Width = 400
        wykres.Parent.Height = 300
        
        ' Legenda
        Set series = wykres.SeriesCollection(1)
        pointIndex = 1
        With wykres.Legend
            .Position = xlLegendPositionRight
            .Font.Size = 12
            .Font.Bold = True
        End With

        ' Procenty oraz wartosc widoczna na wykresie
        With wykres.FullSeriesCollection(1)
            .HasDataLabels = True
            With .DataLabels
                .ShowValue = False
                .ShowPercentage = True
                .Separator = " "
                .Font.Size = 12
                .Font.Bold = True
            End With
        End With
        
        ' Kolorowanie wykresu wed�ug s�ownika
        For Each odp In sortedOdpowiedzi.keys
            series.Points(pointIndex).Format.Fill.ForeColor.RGB = odpowiedziSlownik(odp)(0)
            pointIndex = pointIndex + 1
        Next odp
        
        ' Przesuni�cie wykresu
        outputRow = rowIndex + 26
    Next i

    MsgBox "Wszystkie wykresy zosta�y wygenerowane na nowym arkuszu!", vbInformation
End Sub
