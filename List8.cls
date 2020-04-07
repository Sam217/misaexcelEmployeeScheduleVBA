VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CalcSchTest1_Click()
Dim params618 As typeParams, params61430 As typeParams, paramsW As typeParams, paramsSat As typeParams, _
    paramsSat2 As typeParams
    
    Dim ret As Integer
    ret = checkGlobalVars()
    If ret > 0 Then
        Exit Sub
    End If
    With params61430
        Set .target = Range(dateString)
            .soffset = 0
            .dayfrom = 1
            .dayto = 5
            .shiftinterval = 0
        Set .lDepend = Nothing
            .noDayOfWeekRepeat = False
            .noDayBefore = False
            .noDayAfter = False
            .wkndRule = 0
            .perWeek = True
            .shift_type = 0
    End With
    
    params618 = params61430
    With params618
        'Set .target = Range("B10") 'Range("D10")
            .soffset = 1
            .shiftinterval = 7
        Set .lDepend = Range(params61430.target.Offset(0, 1).Address, params61430.target.Offset(0, params618.soffset).Address) 'Set .lDepend = Nothing 'Range("C10")
            .noDayOfWeekRepeat = True
            .noDayBefore = False
            .noDayAfter = False
            .perWeek = False
            .shift_type = 1
    End With
    
    paramsW = params618
    With paramsW
        'Set .target = Range("E10")
            .soffset = 2
            .dayfrom = 6
            .dayto = 6 'zde jen "sobota", protože výpoèet probìhne na celý víkend, tedy uplatní se i na nedìli...
            .shiftinterval = 0
        Set .lDepend = Range(params61430.target.Offset(0, 1).Address, params61430.target.Offset(0, paramsW.soffset).Address) 'Set .lDepend = Range("c10:d10")
            .noDayOfWeekRepeat = False
            .wkndRule = 2 '2ka znaèí pøípad, že pravidlo musí být splnìno pro obì služby
            .shift_type = 2
    End With
    
    paramsSat = paramsW
    With paramsSat
        'Set .target = Range("F10")
            .soffset = 3
        Set .lDepend = Range(params61430.target.Offset(0, 1).Address, params61430.target.Offset(0, paramsSat.soffset).Address) 'Set .lDepend = Range("E10:C10")
            .wkndRule = 3
            .shift_type = 3
    End With
    
    paramsSat2 = paramsSat
    With paramsSat2
        'Set .target = Range("G10")
        .soffset = 4
        Set .lDepend = Range(params61430.target.Offset(0, 1).Address, params61430.target.Offset(0, paramsSat2.soffset).Address) 'Set .lDepend = Range("F10:C10")
        .wkndRule = 4
        .shift_type = 4
    End With
    markHolidays
    calcSchedule params61430, Range("A1")  'od 6 do pul 3
    calcSchedule params618, Range("A1")  '12ctky
    calcSchedule paramsW, Range("A1")  'sobota+nedele
    calcSchedule paramsSat, Range("A1")  'jen sobota
    calcSchedule paramsSat2, Range("A1")  'jen sobota prisluzba
End Sub

Private Sub deletebttn1_Click()
    Dim rg As Range
    Dim helprng As Range
    Dim ret As Integer 'check global strings
    ret = checkGlobalVars()
    If ret > 0 Then
        Exit Sub
    End If
    Set helprng = Range(dateString)
    Set dtRng = Range(dateString, helprng.Offset(366, 0).Address)
    Set rng2 = Range(helprng.Offset(-1, 1).Address, helprng.Offset(-1, 5).Address)
    count = 0
    For Each cl In rng2
        If Not IsEmpty(cl) Then
           count = count + 1
        End If
    Next
    Set rg = Range("B2:M7")
    rg.ClearContents
    
    For Each cell In dtRng
        If Range(dateStringFrom).Value2 <= cell.Value2 And _
        Range(dateStringTo).Value2 >= cell.Value2 Then
            For i = 1 To count
                cell.Offset(0, i).clear
            Next
        End If
    Next
End Sub

Private Sub CheckMistakes_Click()
    Dim clear As Boolean
    clear = checkForErrorsAlt(Range("C10"), 7, False)
    clear = checkForErrors2Alt(Range("D10"), 1, clear)
    clear = checkForErrors3Alt(Range("E10"), clear)
End Sub

Private Sub RecalcBttn_Click()
    ReCalc
End Sub

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



Private Sub Worksheet_SelectionChange(ByVal target As Range)
    'MsgBox "selection change"
End Sub

