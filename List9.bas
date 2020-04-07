Private Sub CalculateSchedule_Click()
    Dim tgt As Range
    Dim params618 As typeParams, params61430 As typeParams, paramsW As typeParams, paramsSat As typeParams, _
    paramsSat2 As typeParams
    
    Set params618.target = Range("C10")
    params618.daystart = 0
    params618.numofdays = 5
    params618.shiftInterval = 7
    Set params618.lDepend = Nothing
    params618.noDayOfWeekRepeat = True
    params618.noDayBefore = False
    params618.noDayAfter = False
    params618.wkndRule = False
    
    params61430 = params618
    Set params61430.target = Range("D10")
    Set params61430.lDepend = Range("C10")
    params61430.noDayBefore = True
    params61430.noDayAfter = True
    
    paramsW = params618
    Set paramsW.target = Range("E10")
    paramsW.daystart = 5
    paramsW.numofdays = 6
    paramsW.shiftInterval = 0
    Set paramsW.lDepend = Range("c10:d10")
    paramsW.noDayOfWeekRepeat = False
    paramsW.wkndRule = True
    
    paramsSat = paramsW
    Set paramsSat.target = Range("F10")
    Set paramsSat.lDepend = Range("E10:C10")
    paramsSat.wkndRule = False
    
    paramsSat2 = paramsSat
    Set paramsSat2.target = Range("G10")
    Set paramsSat2.lDepend = Range("F10:C10")
    
    calcSchedule params618 '12ctky
    calcSchedule params61430 'od 6 do pul 3
    calcSchedule paramsW 'sobota+nedele
    calcSchedule paramsSat 'jen sobota
    calcSchedule paramsSat2 'jen sobota prisluzba
    
    'Set tgt = Range("D10")
    'calcSchedule tgt, 0, 5, 7, Range("C10"), True, True, True, False
    'calcSchedule Range("E10"), 5, 6, 0, Range("c10:d10"), False, False, False, True
    'calcSchedule Range("F10"), 5, 6, 0, Range("E10:C10"), False, False, False, False
    'calcSchedule Range("G10"), 5, 6, 0, Range("F10:C10"), False, False, False, False
End Sub

Private Sub NOT_Worksheet_Change(ByVal target As Range)
    'Variables
    Dim KeyCells As Range
    Set KeyCells = Range("A1")
    Dim emps(1 To 3) As Long
    Dim empsMin As Long, minIndex As Long, rndval As Long
    Dim arrIdx(1 To 3) As Long
    Dim contains As Boolean, var As Variant
    
    'Code
    If Not Application.Intersect(KeyCells, Range(target.Address)) Is Nothing _
    Then
        For i = 1 To 3
            rndval = Int(3 * Rnd + 1)
            var = Application.Match(rndval, arrIdx, 0)
            While Not IsError(var)
                rndval = Int(3 * Rnd + 1)
                var = Application.Match(rndval, arrIdx, 0)
            Wend
            arrIdx(i) = rndval
        Next
        
        Dim employeeRowOffset As Long: employeeRowOffset = 0
        'If KeyCells.Value >= 1 Then
           ' emps(1) = 1
        'End If
        For i = 0 To 6
            For j = 0 To 2
                Range("C3").Offset(i * 4 + j, 0).ClearContents
            Next
        Next
        For i = 0 To 6
                If (i Mod 3) = 0 Then
                    minIndex = arrIdx(1)
                'employeeRowOffset = Int(2 * Rnd + 1)
                ElseIf (i Mod 3) = 1 Then
                    'minIndex = Int(2 * Rnd + 2)
                    minIndex = arrIdx(2)
                Else
                    empsMin = WorksheetFunction.min(emps)
                    minIndex = WorksheetFunction.Match(empsMin, emps, 0)
                End If
                emps(minIndex) = emps(minIndex) + 1
                Range("C3").Offset(i * 4 + (minIndex - 1), 0).Value = 1
        Next
    End If
' The variable KeyCells contains the cells that will
    ' cause an alert when they are changed.
    'Set KeyCells = Range(Target.Address)

'If Not Application.Intersect(KeyCells, Range(Target.Address)) _
           Is Nothing Then

' Display a message when one of the designated cells has been
        ' changed.
        ' Place your code here.
    'KeyCells.Cells.Offset(4, 0) = Target.Address

'End If
End Sub

Private Sub CheckForErrorsBttn_Click()
    Dim clear As Boolean
    clear = checkForErrors(Range("C10"), 7, False)
    clear = checkForErrors2(Range("D10"), clear)
    clear = checkForErrors3(Range("E15"), clear)
End Sub

Private Sub Worksheet_SelectionChange(ByVal target As Range)
 Dim KeyCells As Range

' The variable KeyCells contains the cells that will
    ' cause an alert when they are changed.
    Set KeyCells = Range("C3")

If Not Application.Intersect(KeyCells, Range(target.Address)) _
           Is Nothing Then

' Display a message when one of the designated cells has been
        ' changed.
        ' Place your code here.
        'MsgBox "Cell " & Target.Address & " has changed."

End If
End Sub
