Attribute VB_Name = "Module2"
Sub calcScheduleAlt(ByVal target As Range, daystart As Integer, numofdays As Integer, shiftinterval As Integer, _
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
    ReDim arrIdx(empColl.count - 1)
    weeksLookBack = empColl.count \ 5
    
        For j = 0 To 364 'empColl.Count
            target.Offset(j, 0).clear
        Next
        Dim unique As Boolean, balanced As Boolean, multi5 As Boolean, b_swap As Boolean
        Dim emp As clsPerson
        Dim rand As Long, swap As Long
        unique = False
        balanced = True
        multi5 = False
        If empColl.count Mod 5 = 0 Then
            multi5 = True
        End If
        Dim k As Integer, i As Integer
        i = 0
    
    If shiftinterval >= empColl.count Then
    'nyní pojistka pro pøípad, že shiftinterval je nekompatibilní s požadavkem "noDayOfWeekRepeat"
        If noDayOfWeekRepeat Then
            shiftinterval = empColl.count - 2
        Else
            shiftinterval = empColl.count - 1
        End If
        'if count of employees with same minimum is less than shiftInterval, the request of shiftInterval can't be fulfilled
        MsgBox "S aktuálním poètem zamìstnancù je nejdelší možná pauza mezi smìnami " & shiftinterval & " dní." _
        & " Nastavena pauza " & shiftinterval & " dní."
    End If
    For d = 0 To 364 'empColl.Count
        If d Mod 7 < numofdays And d Mod 7 >= daystart Then 'numofdays = do jakého dne pojedeme, daystart = od jakého dne zaèneme
            'here is a test if arrIdx has been assigned to all employees i.e. no one is missing from schedule
            'which can happen if someone has very big shift deficit
            If Not unique And i >= empColl.count And balanced And Not multi5 Then
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
            If i < empColl.count Or Not unique Or Not balanced Then
                Set empsEqual = Nothing
                Set empsEqual = findMinOfClsPersons(empColl)
                'if two or more employees have the same minimal shift count we choose randomly
                'who gets the next shift. If only one has the minimum count, then rand will be always zero
                k = i Mod empColl.count
                Randomize
                'find the next min count
                empsMin = empsEqual(1).count
                'Set empsMinSecond = Nothing
                'Set empsMinSecond = findSecondMinOfClsPersons(empColl, empsMin)
                b_swap = True 'check if one employee doesn't have another shift day after
                balanced = True
                Do While b_swap
                    empsMin = empsEqual(1).count
                    rand = Round((empsEqual.count - 1) * Rnd + 1)
                    minIndex = empsEqual(rand).Id
                    arrIdx(k) = minIndex
                    b_swap = False
                    Dim rowoffset As Integer
                    If d < shiftinterval Then
                        rowoffset = d
                    Else
                        rowoffset = shiftinterval
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
                        For cnt = 0 To lDepend.Columns.count - 1
                            If (lDepend.Cells(1, 1).Offset(d, cnt).Value = empsEqual(rand).Name) Then
                            b_swap = True
                            empsEqual.Remove rand
                            Exit For
                            End If
                        Next
                    End If
                    If Not b_swap And noDayBefore And Not lDepend Is Nothing Then
                        For cnt = 0 To lDepend.Columns.count - 1
                            If (lDepend.Cells(1, 1).Offset(d - 1, cnt).Value = empsEqual(rand).Name) Then
                            b_swap = True
                            empsEqual.Remove rand
                            Exit For
                            End If
                        Next
                    End If
                    If Not b_swap And noDayAfter And Not lDepend Is Nothing Then
                        For cnt = 0 To lDepend.Columns.count - 1
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
                        'If epmsEqual.Count = 2 Then 'check if the other employee with the same minimum count
                            'If (lDepend.Cells(1, 1).Offset(d + 2 + 7, 0).Value = empsEqual(2 - rand).Name) Then
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
                    If b_swap And empsEqual.count = 0 Then
                        Set empsEqual = findSecondMinOfClsPersons(empColl, empsMin)
                        balanced = False 'spoléháme se na shiftinterval < empColl.count. Díky této podmínce se ale mùže stát, že index je takový, že arrIdx
                        'se stane unikátním. Pak nechceme, aby se unikátnost kontrolovala a tedy byla stále false. Cyklus se pak bude stále snažit vybalancovat.
                        If empsEqual(1).count = empsMin Then 'Zde je snaha aby pøi nemožnosti splnìní všech požadavkù program neskonèil v nekoneèném cyklu.
                            minIndex = -1 'make arridx position at i invalid and reassign in the next cycle
                            i = i - 1
                            MsgBox "Chyba: pro øádek " & d & "se nepodaøilo splnit všechny požadavky a nemohl být obsloužen."
                            Exit Do
                        End If
                    End If
                Loop
            Else
                k = i Mod empColl.count
                minIndex = arrIdx(k)
            End If
                
            For Each emp In empColl 'another check for whatever happened above, if requirements are met, if not, decrement d to d - 1 and try again
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
                        For cnt = 0 To lDepend.Columns.count - 1
                            If (lDepend.Cells(1, 1).Offset(d, cnt).Value = emp.Name) Then
                            balanced = False
                            b_swap = False
                            Exit For
                            End If
                        Next
                    End If
                    If balanced And noDayBefore And Not lDepend Is Nothing Then
                        For cnt = 0 To lDepend.Columns.count - 1
                            If (lDepend.Cells(1, 1).Offset(d - 1, cnt).Value = emp.Name) Then
                            balanced = False
                            b_swap = False
                            Exit For
                            End If
                        Next
                    End If
                    If balanced And noDayAfter And Not lDepend Is Nothing Then
                        For cnt = 0 To lDepend.Columns.count - 1
                            If (lDepend.Cells(1, 1).Offset(d + 1, cnt).Value = emp.Name) Then
                            balanced = False
                            b_swap = False
                            Exit For
                            End If
                        Next
                    End If
                    End If
                    If b_swap Then
                        emp.count = emp.count + 1
                        'target here is a cell selected by us, recently "C3"
                        target.Offset(d, 0).Value = emp.Name
                        If i Mod empColl.count = 0 Then
                            target.Offset(d, 0).Font.Bold = True
                        End If
                        emplist.Cells(2, minIndex + 1).Value = emp.count
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
        cRowoffset = lDepend.Columns.count
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
    
    '-----------Date processing part
    Dim datecell As Range, startDt As Range
    Set datecell = Range("B10") 'možná lepší pøedávat parametrem
    Set startDt = Range("I9") 'možná lepší pøedávat parametrem
    Dim mydate As Date, startDate As Date, endDate As Date
    mydate = CDate(datecell.Value2)
    
    Dim enddatestr As String
    Dim dtyear As Integer, daysinyear As Integer
    dtyear = Year(mydate)
    enddatestr = dtyear + 1 & "/1/1"
    endDate = CDate(enddatestr)
    daysinyear = DateDiff("d", mydate, endDate)
    
    startDate = CDate(startDt.Value2)
    endDate = CDate(startDt.Offset(1, 0).Value2)
    
    '------------Date processing part end
    
    Dim emplist As Range
    'here we must specify the cell where table with employees starts, in our test case "O1"
    Set emplist = Range("A1")
    Set empColl = LoadEmployees(emplist) 'employees Collection
    ReDim arrIdx(empColl.count - 1)
    weeksLookBack = empColl.count \ 5 'pomocná proménná - pokud je poèet zamìstnancù násobek 5ti,
    'musíme nìjak zjistit, kolik týdnù zpìt se podívat, jestli tam zamìstnanec nemá službu.
    'Code
        'For j = 0 To daysinyear - 1 'empColl.Count
         '   params.Target.Offset(j, 0).clear
        'Next
        Dim unique As Boolean, balanced As Boolean, multi5 As Boolean, b_swap As Boolean
        Dim emp As clsPerson
        Dim rand As Long, swap As Long
        unique = False
        balanced = True
        multi5 = False
        If empColl.count Mod 5 = 0 Then
            multi5 = True
        End If
        Dim k As Integer, i As Integer, d As Integer
        i = 0
        d = 0
    
    If params.shiftinterval >= empColl.count Then
    'nyní pojistka pro pøípad, že shiftinterval je nekompatibilní s požadavkem "noDayOfWeekRepeat"
        If params.noDayOfWeekRepeat Then
            params.shiftinterval = empColl.count - 2
        Else
            params.shiftinterval = empColl.count - 1
        End If
        'if count of employees with same minimum is less than shiftInterval, the request of shiftInterval can't be fulfilled
        MsgBox "S aktuálním poètem zamìstnancù je nejdelší možná pauza mezi smìnami " & params.shiftinterval & " dní." _
        & " Nastavena pauza " & params.shiftinterval & " dní."
    End If
    Dim wd As Integer, debugCounter As Integer: debugCounter = 0
    Dim generate As Boolean: generate = True
    Do While d < daysinyear And debugCounter < 366
        If mydate >= startDate And mydate <= endDate Then
        params.target.Offset(d, 0).clear
        wd = Weekday(mydate, vbMonday)
        'If d Mod 7 < params.dayto And d Mod 7 >= params.dayfrom Then 'dayto = do jakého dne pojedeme, dayfrom = od jakého dne zaèneme ''STARÉ
        If wd >= params.dayfrom And wd <= params.dayto Then
            'here is a test if arrIdx has been assigned to all employees i.e. no one is missing from schedule
            'which can happen if someone has very big shift deficit
            If Not unique And i >= empColl.count And balanced And Not multi5 Then
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
            If i < empColl.count Or Not unique Or Not balanced Then
                Set empsEqual = Nothing
                Set empsEqual = findMinOfClsPersons(empColl)
                'if two or more employees have the same minimal shift count we choose randomly
                'who gets the next shift. If only one has the minimum count, then rand will be always zero
                k = i Mod empColl.count
                Randomize
                'find the next min count
                empsMin = empsEqual(1).count(0)
                'Set empsMinSecond = Nothing
                'Set empsMinSecond = findSecondMinOfClsPersons(empColl, empsMin)
                b_swap = True 'check if one employee doesn't have another shift day after
                balanced = True
                'If (params.perWeek And wd = 1) Or Not params.perWeek Then : generate = True 'Asi to není potøeba...
                Do While b_swap
                    empsMin = empsEqual(1).count(0)
                    If generate Then: rand = Round((empsEqual.count - 1) * Rnd + 1)
                    generate = True
                    b_swap = False
                    Dim rowoffset As Integer
                    If d < params.shiftinterval Then
                        rowoffset = d
                    Else
                        rowoffset = params.shiftinterval
                    End If
                    If params.perWeek Then 'if calculate schedule is for whole work week, d needs to be shifted by 5 days (mon - fri)
                        d = d + 4 - wd + 1 'important - final assigment of employee to work this shift
                        minIndex = empsEqual(rand).Id
                        arrIdx(k) = minIndex
                        'generate = False
                    Else 'BIG OTHERWISE
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
                        For cnt = 0 To params.lDepend.Columns.count - 1
                            If (params.lDepend.Cells(1, 1).Offset(d, cnt).Value = empsEqual(rand).Name) Then
                            b_swap = True
                            empsEqual.Remove rand
                            Exit For
                            End If
                        Next
                    End If
                    If Not b_swap And params.noDayBefore And Not params.lDepend Is Nothing Then
                        For cnt = 0 To params.lDepend.Columns.count - 1
                            If (params.lDepend.Cells(1, 1).Offset(d - 1, cnt).Value = empsEqual(rand).Name) Then
                            b_swap = True
                            empsEqual.Remove rand
                            Exit For
                            End If
                        Next
                    End If
                    If Not b_swap And params.noDayAfter And Not params.lDepend Is Nothing Then
                        For cnt = 0 To params.lDepend.Columns.count - 1
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
                        If (Not params.perWeek) Then
                            If (params.lDepend.Cells(1, 1).Offset(d + 2, 0).Value = empsEqual(rand).Name) Then
                                b_swap = True
                                empsEqual.Remove rand
                            ElseIf empsEqual.count = 2 Then 'check if the other employee with the same minimum count does not meet the condition for next week (thus d + 2 + 7)
                                If (params.lDepend.Cells(1, 1).Offset(d + 2 + 7, 0).Value = empsEqual(3 - rand).Name) Then '...then try this employe instead of current one if he meets all conditions for this week also..
                                    generate = False
                                    rand = 3 - rand
                                    b_swap = True
                                    'balanced = False
                                End If
                            End If
                        Else
                            If (params.lDepend.Cells(1, 1).Offset(d + 2, 0).Value = empsEqual(rand).Name) Or _
                            (params.lDepend.Cells(1, 1).Offset(d + 2, 1).Value = empsEqual(rand).Name) Then
                                b_swap = True
                                empsEqual.Remove rand
                            ElseIf empsEqual.count = 2 Then 'check if the other employee with the same minimum count does not meet the condition for next week (thus d + 2 + 7)
                                If (params.lDepend.Cells(1, 1).Offset(d + 2 + 7, 0).Value = empsEqual(3 - rand).Name) Or _
                                (params.lDepend.Cells(1, 1).Offset(d + 2 + 7, 1).Value = empsEqual(3 - rand).Name) Then '...then try this employe instead of current one if he meets all conditions for this week also..
                                    generate = False
                                    rand = 3 - rand
                                    b_swap = True
                                    'balanced = False
                                End If
                            End If
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
                    If b_swap And empsEqual.count = 0 Then
                        Set empsEqual = findSecondMinOfClsPersons(empColl, empsMin)
                        balanced = False 'spoléháme se na shiftinterval < empColl.count. Díky této podmínce se ale mùže stát, že index je takový, že arrIdx
                        'se stane unikátním. Pak nechceme, aby se unikátnost kontrolovala a tedy byla stále false. Cyklus se pak bude stále snažit vybalancovat.
                        If empsEqual(1).count(0) = empsMin Then 'Zde je snaha aby pøi nemožnosti splnìní všech požadavkù program neskonèil v nekoneèném cyklu.
                            minIndex = -1 'make arridx position at i invalid and reassign in the next cycle
                            i = i - 1
                            MsgBox "Chyba: pro øádek " & d & "se nepodaøilo splnit všechny požadavky a nemohl být obsloužen."
                            Exit Do
                        End If
                    ElseIf Not b_swap Then 'important - final assigment of employee to work this shift
                        minIndex = empsEqual(rand).Id
                        arrIdx(k) = minIndex
                    End If
                    'ENDIF OF BIG OTHERWISE
                    End If
                Loop
            Else
                If params.perWeek Then 'if calculate schedule is for whole work week, d needs to be shifted by 5 days (mon - fri)
                    d = d + 4 - wd + 1 'important - final assigment of employee to work this shift
                End If
                k = i Mod empColl.count
                minIndex = arrIdx(k)
            End If
            'if unique is true, we need to re-check if conditions are met
            For Each emp In empColl
                If emp.Id = minIndex Then
                b_swap = False 'warning! value was previously inversed, result is now the same, but the meaning wasn't correct before
                    If unique Then
                    For cnt = 1 To rowoffset
                        'check again, if shiftInterval is fullfilled etc.
                        If (params.target.Offset(d - cnt, 0).Value = emp.Name) Then
                            balanced = False
                            b_swap = True
                            unique = False
                            Exit For
                        End If
                    Next
                    'check if this employee doesn't have another shift this day already
                    If balanced And Not params.lDepend Is Nothing Then
                        For cnt = 0 To params.lDepend.Columns.count - 1
                            If (params.lDepend.Cells(1, 1).Offset(d, cnt).Value = emp.Name) Then
                            balanced = False
                            b_swap = True
                            unique = False
                            Exit For
                            End If
                        Next
                    End If
                    If balanced And params.noDayBefore And Not params.lDepend Is Nothing Then
                        For cnt = 0 To params.lDepend.Columns.count - 1
                            If (params.lDepend.Cells(1, 1).Offset(d - 1, cnt).Value = emp.Name) Then
                            balanced = False
                            b_swap = True
                            unique = False
                            Exit For
                            End If
                        Next
                    End If
                    If balanced And params.noDayAfter And Not params.lDepend Is Nothing Then
                        For cnt = 0 To params.lDepend.Columns.count - 1
                            If (params.lDepend.Cells(1, 1).Offset(d + 1, cnt).Value = emp.Name) Then
                            balanced = False
                            b_swap = True
                            unique = False
                            Exit For
                            End If
                        Next
                    End If
                    If (Not b_swap And params.wkndRule And Not params.lDepend Is Nothing) Then
                        If (params.lDepend.Cells(1, 1).Offset(d + 2, 0).Value = emp.Name) Or _
                        (params.lDepend.Cells(1, 1).Offset(d + 2, 1).Value = emp.Name) Then
                            b_swap = True
                            balanced = False
                            unique = False
                        End If
                    End If
                    End If
                    If Not b_swap Then
                        'target here is a cell selected by us, recently "C10"
                        If params.perWeek Then
                            Dim enddays As Integer
                            enddays = 0
                            If d > daysinyear - 1 Then: enddays = d - daysinyear + 1
                            emp.count(0) = emp.count(0) + 5 - wd + 1 - enddays
                            emp.count(params.shift_type + 1) = emp.count(params.shift_type + 1) + 5 - wd + 1 - enddays
                            For Z = d - (4 - wd + 1) To d - enddays
                            params.target.Offset(Z, 0).Value = emp.Name
                            params.target.Offset(Z, 0).Font.Name = "Times New Roman"
                            Next
                            If i Mod empColl.count = 0 Then
                            params.target.Offset(d - 4 - wd + 1, 0).Font.Bold = True
                            End If
                        Else
                            emp.count(0) = emp.count(0) + 1
                            emp.count(params.shift_type + 1) = emp.count(params.shift_type + 1) + 1
                            
                            params.target.Offset(d, 0).Value = emp.Name
                            params.target.Offset(d, 0).Font.Name = "Times New Roman"
                            ' If params.shift_type > 1 Then
                            '     params.Target.Offset(d, 0).Interior.Color = RGB(12, 192, 255)
                            ' End If
                            If i Mod empColl.count = 0 Then
                            params.target.Offset(d, 0).Font.Bold = True
                            End If
                        End If
                        emplist.Cells(2, minIndex + 1).Value = emp.count(0)
                        i = i + 1
                    Else
                        d = d - 1
                        debugCounter = debugCounter + 1
                        If debugCounter = 366 Then
                            MsgBox "Selhání pøi pokusu o výpoèet služby v poøadí " & d + 1
                        End If
                    End If
                    Exit For
                End If
            Next
        End If
        End If
        'POST PROCESSING!!! (MOVE to separate procedure when possible)
        If wd > 5 Then
            'params.Target.Offset(d, 0).Interior.Color = RGB(12, 192, 255)
            params.target.Offset(d, 0).Interior.ColorIndex = 37
            With params.target.Offset(d, 0).Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With params.target.Offset(d, 0).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With params.target.Offset(d, 0).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With params.target.Offset(d, 0).Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
        End If
        'END POST PROCESSING!!!
        d = d + 1
        mydate = CDate(datecell.Offset(d, 0).Value2)
    Loop
    'Dim cRowoffset As Integer
    'If Not params.lDepend Is Nothing Then
    '    cRowoffset = params.lDepend.Columns.Count
    'Else: cRowoffset = 0
    'End If
    For Each emp In empColl
        'idx = ""
        'For l = 0 To UBound(arrIdx)
        '    If emp.Id = arrIdx(l) Then
        '        If idx = "" Then
        '            idx = l + 1
        '        Else
        '            idx = idx & "," & l + 1
        '        End If
        '    End If
        'Next
        'emplist.Cells(3 + cRowoffset, emp.Id + 1).ClearContents
        'emplist.Cells(3 + cRowoffset, emp.Id + 1).Value = idx
        emplist.Cells(3 + params.shift_type, emp.Id + 1).ClearContents
        emplist.Cells(3 + params.shift_type, emp.Id + 1).Value = emp.count(params.shift_type + 1)
    Next
    'markHolidays
End Sub
Function checkForErrorsAlt(ByVal target As Range, shiftinterval As Integer, errfound As Boolean) As Boolean
    'kontrola pøekrývajících se služeb týž den a kontrola pauzy mezi službami (shiftInterval)
    Dim rg As Range
    Set rg = target.CurrentRegion
    For i = 1 To rg.Rows.count
        For j = 1 To 3 'rg.Columns.Count
            If target.Cells(i, j).Value = target.Cells(i, j + 1).Value And _
            Not IsEmpty(target.Cells(i, j).Value) Then
                target.Cells(i, j).Interior.Color = RGB(255, 32, 32)
                target.Cells(i, j + 1).Interior.Color = RGB(255, 0, 0)
                errfound = True
                target.Cells(i, j).Select
                'MsgBox "Chyba: " & target.Cells(i, j).Value & " má pøekrývající se služby." _
                '& " Øádek: " & i
            ElseIf (target.Cells(i, j).Value = target.Cells(i - shiftinterval, j).Value Or _
            target.Cells(i, j).Value = target.Cells(i + shiftinterval, j).Value) And _
            Not IsEmpty(target.Cells(i, j).Value) Then
                target.Cells(i, j).Interior.Color = RGB(255, 0, 0)
                target.Cells(i - shiftinterval, j).Interior.Color = RGB(192, 127, 0)
                errfound = True
                target.Cells(i, j).Select
                'MsgBox "Chyba: " & target.Cells(i, j).Value & " nesplòuje podmínku alespoò " _
                '& shiftInterval & " mezi službami." & " Øádek: " & i & " ve sloupci " & target.Cells(0, j).Value
            ElseIf Not errfound Then
                'Target.Cells(i, j).Interior.Color = xlNone
                'Target.Cells(i, j + 1).Interior.Color = xlNone
                'Target.Cells(i - shiftInterval, j).Interior.Color = xlNone
            End If
        Next
    Next
    checkForErrorsAlt = errfound
End Function

Function checkForErrors2Alt(ByVal target As Range, shiftinterval As Integer, errfound As Boolean) As Boolean
    'kontrola služeb den pøed a den po (pùvodnì 6-14:30 nemohlo být den pøed nebo den po 6-18h)
    Dim rg As Range
    Set rg = target.CurrentRegion
    For d = 0 To rg.Rows.count
    For cnt = 1 To target.Columns.count
        If Not IsEmpty(target.Cells(1, 1).Offset(d, 0).Value) And (target.Cells(1, 1).Offset(d - shiftinterval, -cnt).Value = target.Cells(1, 1).Offset(d, 0).Value) Then
            target.Cells(1, 1).Offset(d, 0).Interior.Color = RGB(200, 128, 0)
            target.Cells(1, 1).Offset(d - shiftinterval, -cnt).Interior.Color = RGB(200, 128, 0)
            errfound = True
            target.Cells(1, 1).Offset(d, 0).Select
            'MsgBox "Chyba: " & target.Cells(1, 1).Offset(d, 0).Value & " nesplòuje podmínku " & _
            '"služeb po sobì na øádku " & d + 1
            'Exit For
        End If
        If Not IsEmpty(target.Cells(1, 1).Offset(d, 0).Value) And (target.Cells(1, 1).Offset(d + shiftinterval, -cnt).Value = target.Cells(1, 1).Offset(d, 0).Value) Then
            target.Cells(1, 1).Offset(d, 0).Interior.Color = RGB(200, 128, 0)
            target.Cells(1, 1).Offset(d + shiftinterval, -cnt).Interior.Color = RGB(200, 128, 0)
            errfound = True
            target.Cells(1, 1).Offset(d, 0).Select
            'MsgBox "Chyba: " & target.Cells(1, 1).Offset(d, 0).Value & " nesplòuje podmínku " & _
            '"služeb po sobì na øádku " & d + 1
            'Exit For
            End If
        If Not errfound Then
            'Target.Cells(1, 1).Offset(d, 0).Interior.Color = xlNone
            'Target.Cells(1, 1).Offset(d - 1, -cnt).Interior.Color = xlNone
            'Target.Cells(1, 1).Offset(d + 1, -cnt).Interior.Color = xlNone
        End If
    Next
    Next
    checkForErrors2Alts = errfound
End Function

Function checkForErrors4Alt(ByVal target As Range, shiftinterval As Integer, errfound As Boolean) As Boolean
    'kontrola služeb den pøed a den po (pùvodnì 6-14:30 nemohlo být den pøed nebo den po 6-18h)
    Dim rg As Range
    Set rg = target.CurrentRegion
    For d = 0 To 366 'rg.Rows.count
    For cnt = 0 To 2 'target.Columns.count
        If Not IsEmpty(target.Cells(1, 1).Offset(d, 0).Value) And (target.Cells(1, 1).Offset(d - shiftinterval, -cnt).Value = target.Cells(1, 1).Offset(d, 0).Value) Then
            target.Cells(1, 1).Offset(d, 0).Interior.Color = RGB(200, 128, 0)
            target.Cells(1, 1).Offset(d - shiftinterval, -cnt).Interior.Color = RGB(200, 128, 0)
            errfound = True
            target.Cells(1, 1).Offset(d, 0).Select
            'MsgBox "Chyba: " & target.Cells(1, 1).Offset(d, 0).Value & " nesplòuje podmínku " & _
            '"služeb po sobì na øádku " & d + 1
            'Exit For
        End If
        If Not IsEmpty(target.Cells(1, 1).Offset(d, 0).Value) And (target.Cells(1, 1).Offset(d + shiftinterval, -cnt).Value = target.Cells(1, 1).Offset(d, 0).Value) Then
            target.Cells(1, 1).Offset(d, 0).Interior.Color = RGB(200, 128, 0)
            target.Cells(1, 1).Offset(d + shiftinterval, -cnt).Interior.Color = RGB(200, 128, 0)
            errfound = True
            target.Cells(1, 1).Offset(d, 0).Select
            'MsgBox "Chyba: " & target.Cells(1, 1).Offset(d, 0).Value & " nesplòuje podmínku " & _
            '"služeb po sobì na øádku " & d + 1
            'Exit For
            End If
        If Not errfound Then
            'Target.Cells(1, 1).Offset(d, 0).Interior.Color = xlNone
            'Target.Cells(1, 1).Offset(d - 1, -cnt).Interior.Color = xlNone
            'Target.Cells(1, 1).Offset(d + 1, -cnt).Interior.Color = xlNone
        End If
    Next
    Next
    checkForErrors2Alts = errfound
End Function

Function checkForErrors3Alt(ByVal target As Range, errfound As Boolean) As Boolean
    'kontrola víkendu (služba o víkendu nemùže být, pokud v pondìlí je naplánována služba od 6:00)
    'cyklus jede po každém øádku, ale mimo víkend jsou prázdné...
    Dim rg As Range
    Set rg = target.CurrentRegion
    For d = 0 To rg.Rows.count
    For cnt = 1 To 2 'Target.Columns.Count
        If Not IsEmpty(target.Cells(1, 1).Offset(d, 0).Value) And (target.Cells(1, 1).Offset(d + 2, -cnt).Value = target.Cells(1, 1).Offset(d, 0).Value) Then
            target.Cells(1, 1).Offset(d, 0).Interior.Color = RGB(200, 128, 0)
            target.Cells(1, 1).Offset(d + 2, -cnt).Interior.Color = RGB(200, 128, 0)
            errfound = True
            target.Cells(1, 1).Offset(d, 0).Select
            'MsgBox "Chyba: " & target.Cells(1, 1).Offset(d, 0).Value & " nesplòuje podmínku " & _
            '"služeb po sobì na øádku " & d + 1
            'Exit For
        End If
        If Not errfound Then
            'Target.Cells(1, 1).Offset(d, 0).Interior.Color = xlNone
            'Target.Cells(1, 1).Offset(d + 2, -cnt).Interior.Color = xlNone
        End If
    Next
    Next
    checkForErrors3Alt = errfound
End Function

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

