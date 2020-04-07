Private Sub Worksheet_Change(ByVal target As Range)
    Dim hlpString As String, hlpString2 As String
    If IsDate(target.Value) Then
        hlpString = LCase(target.Offset(0, -1).Value)
        If hlpString = "od:" Then
            dateStringFrom = target.Address
            hlpString2 = LCase(target.Offset(1, -1).Value)
            If hlpString2 = "do:" Then
                dateStringTo = target.Offset(1, 0).Address
            End If
        ElseIf hlpString = "do:" Then
            dateStringTo = target.Address
            hlpString2 = LCase(target.Offset(-1, -1).Value)
            If hlpString2 = "od:" Then
                dateStringFrom = target.Offset(-1, 0).Address
            End If
        Else
            CalcDates target
        End If
    End If
End Sub


