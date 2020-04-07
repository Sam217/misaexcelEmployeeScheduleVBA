Sub markHolidays()
    Dim rng As Range, rng2 As Range, dtRng As Range, dt As Date, dt2 As Date, count As Integer
    Dim helprng As Range
    
    Set helprng = Range(dateString)
    Set dtRng = Range(dateString, helprng.Offset(366, 0).Address)
    Set rng2 = Range(helprng.Offset(-1, 1).Address, helprng.Offset(-1, 5).Address)
    count = 0
    For Each cl In rng2
        If Not IsEmpty(cl) Then
           count = count + 1
        End If
    Next
    For Each cell In dtRng
        'Set rng = Sheets("List10").UsedRange.Find(cell.Value2)
        'dt = CDate(cell.Value2)
        Set rng = List10.UsedRange
        For Each cl In rng
            If cl.Value2 = cell.Value2 Then
            For i = 0 To count
                cell.Offset(0, i).Interior.ColorIndex = 38
            Next
            End If
        Next
        'Set rng2 = List10.UsedRange.Find(cell.Value2)
        'If it is found put its value on the destination sheet
        ' If Not rng Is Nothing Or Not rng2 Is Nothing Then
        '     For i = 0 To 5
        '         If Not rng Is Nothing Then
        '             If cell.Value2 = rng.Value2 Then
        '             cell.Offset(0, i).Interior.ColorIndex = 38
        '             End If
        '         ElseIf Not rng2 Is Nothing Then
        '             If CDate(cell.Value2) = CDate(rng2.Value2) Then
        '             cell.Offset(0, i).Interior.ColorIndex = 38
        '             End If
        '         End If
        '     Next
        ' End If
    Next
End Sub

