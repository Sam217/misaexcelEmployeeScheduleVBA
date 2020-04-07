Sub calcSchedulePerWeek(ByRef params As typeParams)
    'Variables
    'emps = employees
    Dim empsMin As Long, minIndex As Long, rndval As Long, weeksLookBack As Integer
    Dim arrIdx() As Long
    Dim var As Variant, idx As Variant
    Dim empColl As Collection, empsEqual As New Collection
    
    '-----------Date processing part
    Dim datecell As Range
    Set datecell = Range("B10")
    Dim mydate As Date, enddate As Date
    mydate = CDate(datecell.Value2)
    
    Dim enddatestr As String
    Dim dtyear As Integer, daysinyear As Integer
    dtyear = Year(mydate)
    enddatestr = dtyear + 1 & "/1/1"
    enddate = CDate(enddatestr)
    daysinyear = DateDiff("d", mydate, enddate)
    
    '------------Date processing part end
    
    Dim emplist As Range
    'here we must specify the cell where table with employees starts, in our test case "O1"
    Set emplist = Range("A1")
    Set empColl = LoadEmployees(emplist) 'employees Collection
    ReDim arrIdx(empColl.Count - 1)
    weeksLookBack = empColl.Count \ 5 'pomocná proménná - pokud je počet zaměstnanců násobek 5ti,
    'musíme nějak zjistit, kolik týdnů zpět se podívat, jestli tam zaměstnanec nemá službu.
    'Code
        For j = 0 To daysinyear - 1 'empColl.Count
            params.Target.Offset(j, 0).clear
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
    Dim wd As Integer
    Dim generate As Boolean: generate = True
    Do While d < daysinyear
        wd = Weekday(mydate, vbMonday)
        'If d Mod 7 < params.dayto And d Mod 7 >= params.dayfrom Then 'dayto = do jakého dne pojedeme, dayfrom = od jakého dne začneme ''STARÉ
        If wd >= params.dayfrom And wd <= params.dayto Then
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
                empsMin = empsEqual(1).Count(0)
                'Set empsMinSecond = Nothing
                'Set empsMinSecond = findSecondMinOfClsPersons(empColl, empsMin)
                b_swap = True 'check if one employee doesn't have another shift day after
                balanced = True
                'If (params.perWeek And wd = 1) Or Not params.perWeek Then : generate = True 'Asi to není potřeba...
                Do While b_swap
                    empsMin = empsEqual(1).Count(0)
                    If generate Then: rand = Round((empsEqual.Count - 1) * Rnd + 1)
                    generate = True
                    b_swap = False
                    Dim rowoffset As Integer
                    If d < params.shiftInterval Then
                        rowoffset = d
                    Else
                        rowoffset = params.shiftInterval
                    End If
                    If params.perWeek Then 'if calculate schedule is for whole work week, d needs to be shifted by 5 days (mon - fri)
                        d = d + 4 - wd + 1 'important - final assigment of employee to work this shift
                        minIndex = empsEqual(rand).Id
                        arrIdx(k) = minIndex
                        'generate = False
                    Else 'BIG OTHERWISE
                    'if shiftinterval is not fullfilled, try another employee
                    For cnt = 1 To rowoffset
                        If (params.Target.Offset(d - cnt, 0).Value = empsEqual(rand).Name) Then
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
                        ElseIf empsEqual.Count = 2 Then 'check if the other employee with the same minimum count does not meet the condition for next week (thus d + 2 + 7)
                            If (params.lDepend.Cells(1, 1).Offset(d + 2 + 7, 0).Value = empsEqual(3 - rand).Name) Or _
                            (params.lDepend.Cells(1, 1).Offset(d + 2 + 7, 1).Value = empsEqual(3 - rand).Name) Then '...then try this employe instead of current one if he meets all conditions for this week also..
                                generate = False
                                rand = 3 - rand
                                b_swap = True
                                'balanced = False
                            End If
                        End If
                    End If
                    'check the noDayOfWeelRepeat rule - employee must not work the shift always in the same day of the week
                    'the "balanced" rule solves the case, where number of employees is a multiply of 5 - the schedule then must be random
                    If (Not b_swap And params.noDayOfWeekRepeat And (d >= weeksLookBack * 7)) Then
                        If (params.Target.Offset(d - (weeksLookBack * 7), 0).Value = empsEqual(rand).Name) Then
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
                        If empsEqual(1).Count(0) = empsMin Then 'Zde je snaha aby při nemožnosti splnění všech požadavků program neskončil v nekonečném cyklu.
                            minIndex = -1 'make arridx position at i invalid and reassign in the next cycle
                            i = i - 1
                            MsgBox "Chyba: pro řádek " & d & "se nepodařilo splnit všechny požadavky a nemohl být obsloužen."
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
                k = i Mod empColl.Count
                minIndex = arrIdx(k)
            End If
                
            For Each emp In empColl
                If emp.Id = minIndex Then
                b_swap = False 'warning! value was previously inversed, result is now the same, but the meaning wasn't correct before
                    If unique Then
                    For cnt = 1 To rowoffset
                        'check again, if shiftInterval is fullfilled etc.
                        If (params.Target.Offset(d - cnt, 0).Value = emp.Name) Then
                            balanced = False
                            b_swap = True
                            Exit For
                        End If
                    Next
                    'check if this employee doesn't have another shift this day already
                    If balanced And Not params.lDepend Is Nothing Then
                        For cnt = 0 To params.lDepend.Columns.Count - 1
                            If (params.lDepend.Cells(1, 1).Offset(d, cnt).Value = emp.Name) Then
                            balanced = False
                            b_swap = True
                            Exit For
                            End If
                        Next
                    End If
                    If balanced And params.noDayBefore And Not params.lDepend Is Nothing Then
                        For cnt = 0 To params.lDepend.Columns.Count - 1
                            If (params.lDepend.Cells(1, 1).Offset(d - 1, cnt).Value = emp.Name) Then
                            balanced = False
                            b_swap = True
                            Exit For
                            End If
                        Next
                    End If
                    If balanced And params.noDayAfter And Not params.lDepend Is Nothing Then
                        For cnt = 0 To params.lDepend.Columns.Count - 1
                            If (params.lDepend.Cells(1, 1).Offset(d + 1, cnt).Value = emp.Name) Then
                            balanced = False
                            b_swap = True
                            Exit For
                            End If
                        Next
                    End If
                    If (Not b_swap And params.wkndRule And Not params.lDepend Is Nothing) Then
                        If (params.lDepend.Cells(1, 1).Offset(d + 2, 0).Value = emp.Name) Or _
                        (params.lDepend.Cells(1, 1).Offset(d + 2, 1).Value = emp.Name) Then
                            b_swap = True
                            balanced = False
                        End If
                    End If
                    End If
                    If Not b_swap Then
                        'target here is a cell selected by us, recently "C10"
                        If params.perWeek Then
                            Dim enddays As Integer
                            enddays = 0
                            If d > daysinyear - 1 Then: enddays = d - daysinyear + 1
                            emp.Count(0) = emp.Count(0) + 5 - wd + 1 - enddays
                            emp.Count(params.shift_type + 1) = emp.Count(params.shift_type + 1) + 5 - wd + 1 - enddays
                            For Z = d - (4 - wd + 1) To d - enddays
                            params.Target.Offset(Z, 0).Value = emp.Name
                            params.Target.Offset(Z, 0).Font.Name = "Times New Roman"
                            Next
                            If i Mod empColl.Count = 0 Then
                            params.Target.Offset(d - 4 - wd + 1, 0).Font.Bold = True
                            End If
                        Else
                            emp.Count(0) = emp.Count(0) + 1
                            emp.Count(params.shift_type + 1) = emp.Count(params.shift_type + 1) + 1
                            
                            params.Target.Offset(d, 0).Value = emp.Name
                            params.Target.Offset(d, 0).Font.Name = "Times New Roman"
                            ' If params.shift_type > 1 Then
                            '     params.Target.Offset(d, 0).Interior.Color = RGB(12, 192, 255)
                            ' End If
                            If i Mod empColl.Count = 0 Then
                            params.Target.Offset(d, 0).Font.Bold = True
                            End If
                        End If
                        emplist.Cells(2, minIndex + 1).Value = emp.Count(0)
                        i = i + 1
                    Else
                        d = d - 1
                    End If
                    Exit For
                End If
            Next
        End If
        If wd > 5 Then
            'params.Target.Offset(d, 0).Interior.Color = RGB(12, 192, 255)
            params.Target.Offset(d, 0).Interior.ColorIndex = 37
            With params.Target.Offset(d, 0).Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With params.Target.Offset(d, 0).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With params.Target.Offset(d, 0).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With params.Target.Offset(d, 0).Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
        End If
        d = d + 1
        mydate = CDate(datecell.Offset(d,0).Value2)
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
        emplist.Cells(3 + params.shift_type, emp.Id + 1).Value = emp.Count(params.shift_type + 1)
    Next
End Sub
