Sub calcScheduleAlt(ByVal target As Range, daystart As Integer, numofdays As Integer, shiftInterval As Integer, _
lDepend As Range, noDayOfWeekRepeat As Boolean, noDayBefore As Boolean, noDayAfter As Boolean, wkndRule As Boolean)
    'Variables
    'emps = employees
    Dim empsMin As Long, minIndex As Long, rndval As Long, weeksLookBack As Integer
    Dim arrIdx() As Long
    Dim var As Variant, idx As Variant
    Dim empColl As Collection, empsEqual As New Collection
    
    Dim emplist As Range
    'here we must specify the cell where table with employees starts, in our test case "A1"
    Set emplist = Range("A1")
    Set empColl = LoadEmployees(emplist) 'employees Collection
    ReDim arrIdx(empColl.Count - 1)
    weeksLookBack = empColl.Count \ 5
    
        For j = 0 To 364 'empColl.Count
            target.Offset(j, 0).clear
        Next
        Dim unique As Boolean, balanced As Boolean, multi5 As Boolean, b_swap As Boolean
        Dim emp As clsPerson
        Dim rand As Long, swap As Long
        unique = False
        balanced = True
        multi5 = False
        If empColl.Count Mod 5 = 0 Then
            multi5 = True
        End If
        Dim k As Integer, i As Integer
        i = 0
    
    If shiftInterval >= empColl.Count Then
    'nyní pojistka pro případ, že shiftinterval je nekompatibilní s požadavkem "noDayOfWeekRepeat"
        If noDayOfWeekRepeat Then
            shiftInterval = empColl.Count - 2
        Else
            shiftInterval = empColl.Count - 1
        End If
        'if count of employees with same minimum is less than shiftInterval, the request of shiftInterval can't be fulfilled
        MsgBox "S aktuálním počtem zaměstnanců je nejdelší možná pauza mezi směnami " & shiftInterval & " dní." _
        & " Nastavena pauza " & shiftInterval & " dní."
    End If
    For d = 0 To 364 'empColl.Count
        If d Mod 7 < numofdays And d Mod 7 >= daystart Then 'numofdays = do jakého dne pojedeme, daystart = od jakého dne začneme
            'here is a test if arrIdx has been assigned to all employees i.e. no one is missing from schedule
            'which can happen if someone has very big shift deficit
            If Not unique And i >= empColl.Count And balanced And Not multi5 Then
                unique = True
                For Each emp In empColl
                    For k = 0 To UBound(arrIdx)
                        If emp.Id = arrIdx(k) Then
                         Exit For
                        ElseIf k = UBound(arrIdx) Then
                         unique = False
                        End If
                    Next
                    If Not unique Then
                        Exit For
                    End If
                Next
            End If
            If i < empColl.Count Or Not unique Or Not balanced Then
                Set empsEqual = Nothing
                Set empsEqual = findMinOfClsPersons(empColl)
                'if two or more employees have the same minimal shift count we choose randomly
                'who gets the next shift. If only one has the minimum count, then rand will be always zero
                k = i Mod empColl.Count
                Randomize
                'find the next min count
                empsMin = empsEqual(1).Count
                'Set empsMinSecond = Nothing
                'Set empsMinSecond = findSecondMinOfClsPersons(empColl, empsMin)
                b_swap = True 'check if one employee doesn't have another shift day after
                balanced = True
                Do While b_swap
                    empsMin = empsEqual(1).Count
                    rand = Round((empsEqual.Count - 1) * Rnd + 1)
                    minIndex = empsEqual(rand).Id
                    arrIdx(k) = minIndex
                    b_swap = False
                    Dim rowoffset As Integer
                    If d < shiftInterval Then
                        rowoffset = d
                    Else
                        rowoffset = shiftInterval
                    End If
                    'if shiftinterval is not fullfilled, try another employee
                    For cnt = 1 To rowoffset
                        If (target.Offset(d - cnt, 0).Value = empsEqual(rand).Name) Then
                            b_swap = True
                            'to avoid (despite very low probability) of getting repeatedly the same index from "rand"
                            'for very long time, we could insert here "empsEqual.Remove rand"
                            'after that it is guaranteed to get a different employee id
                            empsEqual.Remove rand
                            Exit For
                        End If
                    Next
                    'check if this employee doesn't have another shift this day already
                    If Not b_swap And Not lDepend Is Nothing Then
                        For cnt = 0 To lDepend.Columns.Count - 1
                            If (lDepend.Cells(1, 1).Offset(d, cnt).Value = empsEqual(rand).Name) Then
                            b_swap = True
                            empsEqual.Remove rand
                            Exit For
                            End If
                        Next
                    End If
                    If Not b_swap And noDayBefore And Not lDepend Is Nothing Then
                        For cnt = 0 To lDepend.Columns.Count - 1
                            If (lDepend.Cells(1, 1).Offset(d - 1, cnt).Value = empsEqual(rand).Name) Then
                            b_swap = True
                            empsEqual.Remove rand
                            Exit For
                            End If
                        Next
                    End If
                    If Not b_swap And noDayAfter And Not lDepend Is Nothing Then
                        For cnt = 0 To lDepend.Columns.Count - 1
                            If (lDepend.Cells(1, 1).Offset(d + 1, cnt).Value = empsEqual(rand).Name) Then
                            b_swap = True
                            empsEqual.Remove rand
                            Exit For
                            End If
                        Next
                    End If
                    'check the weekend rule - employee must not work 12h shift after weekend shift
                    'resp. must not work weekend shift right before 12h shift on monday
                    If (Not b_swap And wkndRule And Not lDepend Is Nothing) Then
                        If (lDepend.Cells(1, 1).Offset(d + 2, 0).Value = empsEqual(rand).Name) Then
                            b_swap = True
                            empsEqual.Remove rand
                        End If
                    End If
                    'check the noDayOfWeelRepeat rule - employee must not work the shift always in the same day of the week
                    'the "balanced" rule solves the case, where number of employees is a multiply of 5 - the schedule then must be random
                    If (Not b_swap And noDayOfWeekRepeat And (d >= weeksLookBack * 7)) Then
                        If (target.Offset(d - (weeksLookBack * 7), 0).Value = empsEqual(rand).Name) Then
                            b_swap = True
                            balanced = False
                            empsEqual.Remove rand
                        End If
                    End If
                    'if employees with min are depleted (didn't match the requirements above), pick another set of employees find next lowest shift count
                    If b_swap And empsEqual.Count = 0 Then
                        Set empsEqual = findSecondMinOfClsPersons(empColl, empsMin)
                        balanced = False 'spoléháme se na shiftinterval < empColl.count. Díky této podmínce se ale může stát, že index je takový, že arrIdx
                        'se stane unikátním. Pak nechceme, aby se unikátnost kontrolovala a tedy byla stále false. Cyklus se pak bude stále snažit vybalancovat.
                        If empsEqual(1).Count = empsMin Then 'Zde je snaha aby při nemožnosti splnění všech požadavků program neskončil v nekonečném cyklu.
                            minIndex = -1 'make arridx position at i invalid and reassign in the next cycle
                            i = i - 1
                            MsgBox "Chyba: pro řádek " & d & "se nepodařilo splnit všechny požadavky a nemohl být obsloužen."
                            Exit Do
                        End If
                    End If
                Loop
            Else
                k = i Mod empColl.Count
                minIndex = arrIdx(k)
            End If
                
            For Each emp In empColl
                If emp.Id = minIndex Then
                b_swap = True
                    If unique Then
                    For cnt = 1 To rowoffset
                        'check again, if shiftInterval is fullfilled etc.
                        If (target.Offset(d - cnt, 0).Value = emp.Name) Then
                            balanced = False
                            b_swap = False
                            Exit For
                        End If
                    Next
                    'check if this employee doesn't have another shift this day already
                    If balanced And Not lDepend Is Nothing Then
                        For cnt = 0 To lDepend.Columns.Count - 1
                            If (lDepend.Cells(1, 1).Offset(d, cnt).Value = emp.Name) Then
                            balanced = False
                            b_swap = False
                            Exit For
                            End If
                        Next
                    End If
                    If balanced And noDayBefore And Not lDepend Is Nothing Then
                        For cnt = 0 To lDepend.Columns.Count - 1
                            If (lDepend.Cells(1, 1).Offset(d - 1, cnt).Value = emp.Name) Then
                            balanced = False
                            b_swap = False
                            Exit For
                            End If
                        Next
                    End If
                    If balanced And noDayAfter And Not lDepend Is Nothing Then
                        For cnt = 0 To lDepend.Columns.Count - 1
                            If (lDepend.Cells(1, 1).Offset(d + 1, cnt).Value = emp.Name) Then
                            balanced = False
                            b_swap = False
                            Exit For
                            End If
                        Next
                    End If
                    End If
                    If b_swap Then
                        emp.Count = emp.Count + 1
                        'target here is a cell selected by us, recently "C3"
                        target.Offset(d, 0).Value = emp.Name
                        If i Mod empColl.Count = 0 Then
                            target.Offset(d, 0).Font.Bold = True
                        End If
                        emplist.Cells(2, minIndex + 1).Value = emp.Count
                        i = i + 1
                    Else
                        d = d - 1
                    End If
                    Exit For
                End If
            Next
        End If
    Next
    Dim cRowoffset As Integer
    If Not lDepend Is Nothing Then
        cRowoffset = lDepend.Columns.Count
    Else: cRowoffset = 0
    End If
    For Each emp In empColl
        idx = ""
        For l = 0 To UBound(arrIdx)
            If emp.Id = arrIdx(l) Then
                If idx = "" Then
                    idx = l + 1
                Else
                    idx = idx & "," & l + 1
                End If
            End If
        Next
        emplist.Cells(3 + cRowoffset, emp.Id + 1).ClearContents
        emplist.Cells(3 + cRowoffset, emp.Id + 1).Value = idx
    Next
End Sub

Sub calcSchedulePerWeek(ByRef params As typeParams)
    'Variables
    'emps = employees
    Dim empsMin As Long, minIndex As Long, rndval As Long, weeksLookBack As Integer
    Dim arrIdx() As Long
    Dim var As Variant, idx As Variant
    Dim empColl As Collection, empsEqual As New Collection
    
    Dim emplist As Range
    'here we must specify the cell where table with employees starts, in our test case "O1"
    Set emplist = Range("A1")
    Set empColl = LoadEmployees(emplist) 'employees Collection
    ReDim arrIdx(empColl.Count - 1)
    weeksLookBack = empColl.Count \ 5 'pomocná proménná - pokud je počet zaměstnanců násobek 5ti,
    'musíme nějak zjistit, kolik týdnů zpět se podívat, jestli tam zaměstnanec nemá službu.
    'Code
        For j = 0 To 364 'empColl.Count
            params.target.Offset(j, 0).clear
        Next
        Dim unique As Boolean, balanced As Boolean, multi5 As Boolean, b_swap As Boolean
        Dim emp As clsPerson
        Dim rand As Long, swap As Long
        unique = False
        balanced = True
        multi5 = False
        If empColl.Count Mod 5 = 0 Then
            multi5 = True
        End If
        Dim k As Integer, i As Integer, d As Integer
        i = 0
        d = 0
    
    If params.shiftInterval >= empColl.Count Then
    'nyní pojistka pro případ, že shiftinterval je nekompatibilní s požadavkem "noDayOfWeekRepeat"
        If params.noDayOfWeekRepeat Then
            params.shiftInterval = empColl.Count - 2
        Else
            params.shiftInterval = empColl.Count - 1
        End If
        'if count of employees with same minimum is less than shiftInterval, the request of shiftInterval can't be fulfilled
        MsgBox "S aktuálním počtem zaměstnanců je nejdelší možná pauza mezi směnami " & params.shiftInterval & " dní." _
        & " Nastavena pauza " & params.shiftInterval & " dní."
    End If
    Do While d < 365
        If d Mod 7 < params.numofdays And d Mod 7 >= params.daystart Then 'numofdays = do jakého dne pojedeme, daystart = od jakého dne začneme
            'here is a test if arrIdx has been assigned to all employees i.e. no one is missing from schedule
            'which can happen if someone has very big shift deficit
            If Not unique And i >= empColl.Count And balanced And Not multi5 Then
                unique = True
                For Each emp In empColl
                    For k = 0 To UBound(arrIdx)
                        If emp.Id = arrIdx(k) Then
                         Exit For
                        ElseIf k = UBound(arrIdx) Then
                         unique = False
                        End If
                    Next
                    If Not unique Then
                        Exit For
                    End If
                Next
            End If
            If i < empColl.Count Or Not unique Or Not balanced Then
                Set empsEqual = Nothing
                Set empsEqual = findMinOfClsPersons(empColl)
                'if two or more employees have the same minimal shift count we choose randomly
                'who gets the next shift. If only one has the minimum count, then rand will be always zero
                k = i Mod empColl.Count
                Randomize
                'find the next min count
                empsMin = empsEqual(1).Count
                'Set empsMinSecond = Nothing
                'Set empsMinSecond = findSecondMinOfClsPersons(empColl, empsMin)
                b_swap = True 'check if one employee doesn't have another shift day after
                balanced = True
                Do While b_swap
                    empsMin = empsEqual(1).Count
                    rand = Round((empsEqual.Count - 1) * Rnd + 1)
                    minIndex = empsEqual(rand).Id
                    arrIdx(k) = minIndex
                    b_swap = False
                    Dim rowoffset As Integer
                    If d < params.shiftInterval Then
                        rowoffset = d
                    Else
                        rowoffset = params.shiftInterval
                    End If
                    If params.perWeek Then
                        d = d + 4
                    Else
                    'if shiftinterval is not fullfilled, try another employee
                    For cnt = 1 To rowoffset
                        If (params.target.Offset(d - cnt, 0).Value = empsEqual(rand).Name) Then
                            b_swap = True
                            'to avoid (despite very low probability) of getting repeatedly the same index from "rand"
                            'for very long time, we could insert here "empsEqual.Remove rand"
                            'after that it is guaranteed to get a different employee id
                            empsEqual.Remove rand
                            Exit For
                        End If
                    Next
                    'check if this employee doesn't have another shift this day already
                    If Not b_swap And Not params.lDepend Is Nothing Then
                        For cnt = 0 To params.lDepend.Columns.Count - 1
                            If (params.lDepend.Cells(1, 1).Offset(d, cnt).Value = empsEqual(rand).Name) Then
                            b_swap = True
                            empsEqual.Remove rand
                            Exit For
                            End If
                        Next
                    End If
                    If Not b_swap And params.noDayBefore And Not params.lDepend Is Nothing Then
                        For cnt = 0 To params.lDepend.Columns.Count - 1
                            If (params.lDepend.Cells(1, 1).Offset(d - 1, cnt).Value = empsEqual(rand).Name) Then
                            b_swap = True
                            empsEqual.Remove rand
                            Exit For
                            End If
                        Next
                    End If
                    If Not b_swap And params.noDayAfter And Not params.lDepend Is Nothing Then
                        For cnt = 0 To params.lDepend.Columns.Count - 1
                            If (params.lDepend.Cells(1, 1).Offset(d + 1, cnt).Value = empsEqual(rand).Name) Then
                            b_swap = True
                            empsEqual.Remove rand
                            Exit For
                            End If
                        Next
                    End If
                    'check the weekend rule - employee must not work shift from 6:00 after weekend shift
                    'resp. must not work weekend shift right before shift from 6:00 on monday
                    If (Not b_swap And params.wkndRule And Not params.lDepend Is Nothing) Then
                        If (params.lDepend.Cells(1, 1).Offset(d + 2, 0).Value = empsEqual(rand).Name) Or _
                        (params.lDepend.Cells(1, 1).Offset(d + 2, 1).Value = empsEqual(rand).Name) Then
                            b_swap = True
                            empsEqual.Remove rand
                        End If
                    End If
                    'check the noDayOfWeelRepeat rule - employee must not work the shift always in the same day of the week
                    'the "balanced" rule solves the case, where number of employees is a multiply of 5 - the schedule then must be random
                    If (Not b_swap And params.noDayOfWeekRepeat And (d >= weeksLookBack * 7)) Then
                        If (params.target.Offset(d - (weeksLookBack * 7), 0).Value = empsEqual(rand).Name) Then
                            b_swap = True
                            balanced = False
                            empsEqual.Remove rand
                        End If
                    End If
                    'if employees with min are depleted (didn't match the requirements above), pick another set of employees find next lowest shift count
                    If b_swap And empsEqual.Count = 0 Then
                        Set empsEqual = findSecondMinOfClsPersons(empColl, empsMin)
                        balanced = False 'spoléháme se na shiftinterval < empColl.count. Díky této podmínce se ale může stát, že index je takový, že arrIdx
                        'se stane unikátním. Pak nechceme, aby se unikátnost kontrolovala a tedy byla stále false. Cyklus se pak bude stále snažit vybalancovat.
                        If empsEqual(1).Count = empsMin Then 'Zde je snaha aby při nemožnosti splnění všech požadavků program neskončil v nekonečném cyklu.
                            minIndex = -1 'make arridx position at i invalid and reassign in the next cycle
                            i = i - 1
                            MsgBox "Chyba: pro řádek " & d & "se nepodařilo splnit všechny požadavky a nemohl být obsloužen."
                            Exit Do
                        End If
                    End If
                    End If
                Loop
            Else
                If params.perWeek Then
                        d = d + 4
                End If
                k = i Mod empColl.Count
                minIndex = arrIdx(k)
            End If
                
            For Each emp In empColl
                If emp.Id = minIndex Then
                b_swap = True
                    If unique Then
                    For cnt = 1 To rowoffset
                        'check again, if shiftInterval is fullfilled etc.
                        If (params.target.Offset(d - cnt, 0).Value = emp.Name) Then
                            balanced = False
                            b_swap = False
                            Exit For
                        End If
                    Next
                    'check if this employee doesn't have another shift this day already
                    If balanced And Not params.lDepend Is Nothing Then
                        For cnt = 0 To params.lDepend.Columns.Count - 1
                            If (params.lDepend.Cells(1, 1).Offset(d, cnt).Value = emp.Name) Then
                            balanced = False
                            b_swap = False
                            Exit For
                            End If
                        Next
                    End If
                    If balanced And params.noDayBefore And Not params.lDepend Is Nothing Then
                        For cnt = 0 To params.lDepend.Columns.Count - 1
                            If (params.lDepend.Cells(1, 1).Offset(d - 1, cnt).Value = emp.Name) Then
                            balanced = False
                            b_swap = False
                            Exit For
                            End If
                        Next
                    End If
                    If balanced And params.noDayAfter And Not params.lDepend Is Nothing Then
                        For cnt = 0 To params.lDepend.Columns.Count - 1
                            If (params.lDepend.Cells(1, 1).Offset(d + 1, cnt).Value = emp.Name) Then
                            balanced = False
                            b_swap = False
                            Exit For
                            End If
                        Next
                    End If
                    End If
                    If b_swap Then
                        'target here is a cell selected by us, recently "C10"
                        If params.perWeek Then
                            Dim enddays As Integer
                            enddays = 0
                            If d > 364 Then: enddays = d - 364
                            emp.Count = emp.Count + 5 - enddays
                            For Z = d - 4 To d - enddays
                            params.target.Offset(Z, 0).Value = emp.Name
                            Next
                            If i Mod empColl.Count = 0 Then
                            params.target.Offset(d - 4, 0).Font.Bold = True
                            End If
                        Else
                            emp.Count = emp.Count + 1
                            params.target.Offset(d, 0).Value = emp.Name
                            If i Mod empColl.Count = 0 Then
                            params.target.Offset(d, 0).Font.Bold = True
                            End If
                        End If
                        emplist.Cells(2, minIndex + 1).Value = emp.Count
                        i = i + 1
                    Else
                        d = d - 1
                    End If
                    Exit For
                End If
            Next
        End If
        d = d + 1
    Loop
    Dim cRowoffset As Integer
    If Not params.lDepend Is Nothing Then
        cRowoffset = params.lDepend.Columns.Count
    Else: cRowoffset = 0
    End If
    For Each emp In empColl
        idx = ""
        For l = 0 To UBound(arrIdx)
            If emp.Id = arrIdx(l) Then
                If idx = "" Then
                    idx = l + 1
                Else
                    idx = idx & "," & l + 1
                End If
            End If
        Next
        emplist.Cells(3 + cRowoffset, emp.Id + 1).ClearContents
        emplist.Cells(3 + cRowoffset, emp.Id + 1).Value = idx
    Next
End Sub

Function checkForErrorsAlt(ByVal target As Range, shiftInterval As Integer, errfound As Boolean) As Boolean
    'kontrola překrývajících se služeb týž den a kontrola pauzy mezi službami (shiftInterval)
    Dim rg As Range
    Set rg = target.CurrentRegion
    For i = 1 To rg.Rows.Count
        For j = 1 To 2 'rg.Columns.Count
            If target.Cells(i, j).Value = target.Cells(i, j + 1).Value And _
            Not IsEmpty(target.Cells(i, j).Value) Then
                target.Cells(i, j).Interior.Color = RGB(255, 32, 32)
                target.Cells(i, j + 1).Interior.Color = RGB(255, 0, 0)
                errfound = True
                target.Cells(i, j).Select
                'MsgBox "Chyba: " & target.Cells(i, j).Value & " má překrývající se služby." _
                '& " Řádek: " & i
            ElseIf target.Cells(i, j).Value = target.Cells(i - shiftInterval, j).Value And _
            Not IsEmpty(target.Cells(i, j).Value) Then
                target.Cells(i, j).Interior.Color = RGB(255, 0, 0)
                target.Cells(i - shiftInterval, j).Interior.Color = RGB(192, 127, 0)
                errfound = True
                target.Cells(i, j).Select
                'MsgBox "Chyba: " & target.Cells(i, j).Value & " nesplňuje podmínku alespoň " _
                '& shiftInterval & " mezi službami." & " Řádek: " & i & " ve sloupci " & target.Cells(0, j).Value
            ElseIf Not errfound Then
                target.Cells(i, j).Interior.Color = xlNone
                target.Cells(i, j + 1).Interior.Color = xlNone
                target.Cells(i - shiftInterval, j).Interior.Color = xlNone
            End If
        Next
    Next
    checkForErrorsAlt = errfound
End Function

Function checkForErrors2Alt(ByVal target As Range, errfound As Boolean) As Boolean
    'kontrola služeb den před a den po (původně 6-14:30 nemohlo být den před nebo den po 6-18h)
    Dim rg As Range
    Set rg = target.CurrentRegion
    For d = 0 To rg.Rows.Count
    For cnt = 1 To target.Columns.Count
        If Not IsEmpty(target.Cells(1, 1).Offset(d, 0).Value) And (target.Cells(1, 1).Offset(d - 1, -cnt).Value = target.Cells(1, 1).Offset(d, 0).Value) Then
            target.Cells(1, 1).Offset(d, 0).Interior.Color = RGB(200, 128, 0)
            target.Cells(1, 1).Offset(d - 1, -cnt).Interior.Color = RGB(200, 128, 0)
            errfound = True
            target.Cells(1, 1).Offset(d, 0).Select
            'MsgBox "Chyba: " & target.Cells(1, 1).Offset(d, 0).Value & " nesplňuje podmínku " & _
            '"služeb po sobě na řádku " & d + 1
            'Exit For
        End If
        If Not IsEmpty(target.Cells(1, 1).Offset(d, 0).Value) And (target.Cells(1, 1).Offset(d + 1, -cnt).Value = target.Cells(1, 1).Offset(d, 0).Value) Then
            target.Cells(1, 1).Offset(d, 0).Interior.Color = RGB(200, 128, 0)
            target.Cells(1, 1).Offset(d + 1, -cnt).Interior.Color = RGB(200, 128, 0)
            errfound = True
            target.Cells(1, 1).Offset(d, 0).Select
            'MsgBox "Chyba: " & target.Cells(1, 1).Offset(d, 0).Value & " nesplňuje podmínku " & _
            '"služeb po sobě na řádku " & d + 1
            'Exit For
            End If
        If Not errfound Then
            target.Cells(1, 1).Offset(d, 0).Interior.Color = xlNone
            target.Cells(1, 1).Offset(d - 1, -cnt).Interior.Color = xlNone
            target.Cells(1, 1).Offset(d + 1, -cnt).Interior.Color = xlNone
        End If
    Next
    Next
    checkForErrors2Alts = errfound
End Function
Function checkForErrors3Alt(ByVal target As Range, errfound As Boolean) As Boolean
    'kontrola víkendu (služba o víkendu nemůže být, pokud v pondělí je naplánována služba od 6:00)
    'cyklus jede po každém řádku, ale mimo víkend jsou prázdné...
    Dim rg As Range
    Set rg = target.CurrentRegion
    For d = 0 To rg.Rows.Count
    For cnt = 1 To target.Columns.Count
        If Not IsEmpty(target.Cells(1, 1).Offset(d, 0).Value) And (target.Cells(1, 1).Offset(d + 2, -cnt).Value = target.Cells(1, 1).Offset(d, 0).Value) Then
            target.Cells(1, 1).Offset(d, 0).Interior.Color = RGB(200, 128, 0)
            target.Cells(1, 1).Offset(d + 2, -cnt).Interior.Color = RGB(200, 128, 0)
            errfound = True
            target.Cells(1, 1).Offset(d, 0).Select
            'MsgBox "Chyba: " & target.Cells(1, 1).Offset(d, 0).Value & " nesplňuje podmínku " & _
            '"služeb po sobě na řádku " & d + 1
            'Exit For
        End If
        If Not errfound Then
            target.Cells(1, 1).Offset(d, 0).Interior.Color = xlNone
            target.Cells(1, 1).Offset(d + 2, -cnt).Interior.Color = xlNone
        End If
    Next
    Next
    checkForErrors3Alt = errfound
End Function
