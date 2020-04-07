Sub ReCalc()
    Dim KeyCells As Range
    Dim rg As Range
    Set rg = Range("B2:L7")
    rg.ClearContents
    Set KeyCells = Range("B10")
        Dim j As Integer: j = 0
        Dim rng As Range, cnt As Integer
        Set rng = Range("B1:Z1")
        Do While Not IsEmpty(KeyCells.Offset(j, 0))
            For k = 1 To 5
                For Each cell In rng
                    If IsEmpty(cell) Then: Exit For  'pozor, tabulka zaměstnanců musí být souvislá...
                    If KeyCells.Offset(j, k).Value2 = cell.Value2 Then
                        'celkový počet (je na prvním řádku)
                        cnt = cell.Offset(1, 0).Value2
                        cnt = cnt + 1
                        cell.Offset(1, 0).Value2 = cnt
                        'konkrétní počet - nutno zvýšit k o 1
                        If KeyCells.Offset(j, k).Interior.ColorIndex = 38 Then
                            'sth
                        Else
                            cnt = cell.Offset(k + 1, 0).Value2
                            cnt = cnt + 1
                            cell.Offset(k + 1, 0).Value2 = cnt
                        End If
                        Exit For 'Předpokládáme, že v horní tabulce zaměstnance je každý zaměstnanec pouze jednou
                    End If
                    'If IsEmpty(cell) Then : Exit For 'pozor, tabulka zaměstnanců musí být souvislá...
                Next
            Next
            j = j + 1
        Loop
End Sub