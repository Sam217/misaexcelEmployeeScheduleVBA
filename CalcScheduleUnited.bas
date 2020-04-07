Sub calcSchedule(ByRef params As typeParams, emplist As Range)
    'markHolidays
    'Variables
    'emps = employees
    Dim empsMin As Long, minIndex As Long, rndval As Long, weeksLookBack As Integer, noWeeksBefore As Integer
    Dim arrIdx() As Long
    Dim var As Variant, idx As Variant
    Dim empColl As Collection, empsEqual As New Collection
    
    '-----------Date processing part
    'Dim datecell As Range, startDt As Range
    'Set datecell = Range("B10") 'možná lepší předávat parametrem
    'Set startDt = Range("I9") 'možná lepší předávat parametrem
    Dim mydate As Date, startDate As Date, endDate As Date
    'mydate = CDate(datecell.Value2)
    mydate = CDate(params.target.Value2)
    Dim targetcells As Range
    Set targetcells = Range(params.target.Offset(0, params.soffset + 1).Address)

    Dim enddatestr As String
    Dim dtyear As Integer, daysinyear As Integer
    dtyear = Year(mydate)
    enddatestr = dtyear + 1 & "/1/1"
    endDate = CDate(enddatestr)
    daysinyear = DateDiff("d", mydate, endDate)
    startDate = CDate(Range(dateStringFrom).Value2)
    endDate = CDate(Range(dateStringTo).Value2)
    
    '------------Date processing part end
    
    'Dim emplist As Range
    'here we must specify the cell where table with employees starts, in our test case "O1"
    'Set emplist = Range("A1")
    Set empColl = LoadEmployees(emplist) 'employees Collection

    If params.perWeek Then
        ReDim arrIdx((empColl.count - 1) * 5)
    Else
        ReDim arrIdx(empColl.count - 1)
    End If
    weeksLookBack = empColl.count \ 5 'pomocná proménná - pokud je počet zaměstnanců násobek 5ti,
    'musíme nějak zjistit, kolik týdnů zpět se podívat, jestli tam zaměstnanec nemá službu.
    'Code
        'For j = 0 To daysinyear - 1 'empColl.Count
         '   targetcells.Offset(j, 0).clear
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
    
    If params.shiftInterval >= empColl.count And params.wkndRule = 0 Then
    'nyní pojistka pro případ, že shiftinterval je nekompatibilní s požadavkem "noDayOfWeekRepeat"
        If params.noDayOfWeekRepeat Then
            params.shiftInterval = empColl.count - 2
        Else
            params.shiftInterval = empColl.count - 1
        End If
        'if count of employees with same minimum is less than shiftInterval, the request of shiftInterval can't be fulfilled
        MsgBox "S aktuálním počtem zaměstnanců je nejdelší možná pauza mezi směnami " & params.shiftInterval & " dní." _
        & " Nastavena pauza " & params.shiftInterval & " dní."
    End If
    Dim wd As Integer, debugCounter As Integer, dbgcntr2 As Integer
    Dim pickAnother As Boolean: pickAnother = True
    Dim findnewmins As Boolean: findnewmins = True
    Dim iwantcontinuebuticant As Boolean: iwantcontinuebuticant = True 'já chci, ne program... není příkaz continue...
    Dim isholiday As Boolean: isholiday = False 'svátky, ne prázdniny
    Dim increment As Boolean: increment = True
    Dim display As Boolean: increment = True
    Dim notPassedWknd1 As Boolean: notPassedWknd1 = False
    noWeeksBefore = 2
    'Dim holcnt as Integer
    Do While d < daysinyear And debugCounter < 366
        increment = True
        wd = Weekday(mydate, vbMonday)
        isholiday = (targetcells.Offset(d, 0).Interior.ColorIndex = 38)
        If Not findnewmins And (((params.perWeek And wd = 1) Or Not params.perWeek)) And (params.wkndRule = 0) Then: findnewmins = True  'Zde by to mělo fungovat správně a zaručit generování když je opravdu potřeba
        If Not findnewmins And (Not isholiday And params.wkndRule > 0 And wd = 5) Then: findnewmins = True 'kontrola - před víkendem chceme generovat nové, aby na sobotě nezůstalo generování vyp.
        'zde se postaráme, aby v létě o prázdninách a o vánocích nebyla služba 6-14:30, v případě této služby od po-pá (perWeek)
        If params.perWeek Then
            If (Month(mydate) >= 7 And Month(mydate) <= 8) Or _
            (Month(mydate) = 12 And Day(mydate) >= 24 And Day(mydate) <= 26) Then 'toto jsou pravidla vynechání služeb o letních prázdninách a vánocích
                iwantcontinuebuticant = False
            Else
                iwantcontinuebuticant = True
            End If
        End If
        
        If mydate >= startDate And mydate <= endDate And iwantcontinuebuticant Then
        'targetcells.Offset(d, 0).clear
        'If d Mod 7 < params.dayto And d Mod 7 >= params.dayfrom Then 'dayto = do jakého dne pojedeme, dayfrom = od jakého dne začneme ''STARÉ 'podmínka rozšířena o kontrolu svátku podle barvy pozadí buňky...
        If (wd >= params.dayfrom And wd <= params.dayto And Not isholiday) Or _
        (wd <> 7 And params.wkndRule > 0 And isholiday) Then
            'pokud je svátek v pondělí, nechceme generovat. Pracují lidé z víkendu.
            If params.wkndRule > 0 And isholiday And wd = 1 Then
                findnewmins = False
            End If
            If isholiday Then
                ' balanced = False
                ' unique = True 'HACK
                If Not balanced Then: unique = False  'tohle je třeba ještě promyslet
            End If
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
                'If Not findnewmins And ((params.perWeek And wd = 1) Or Not params.perWeek) Then : findnewmins = True 'Asi to je potřeba... 'toto tady nefunguje správně, pokud pondělí zrovna "vynecháme z hlavního cyklu"
                If findnewmins And pickAnother Then 'Or empsEqual.count = 0 Then 'toto není dobrý fix...empsEqual by v tomto kroku nikdy neměl být prázdný
                    Set empsEqual = Nothing
                    Set empsEqual = findMinOfClsPersons(empColl, params.shift_type + 1, isholiday)
                    If empsEqual.count = 0 Then
                        MsgBox "chyba"
                    End If
                End If
                dbgcntr2 = 0
                'if two or more employees have the same minimal shift count we choose randomly
                'who gets the next shift. If only one has the minimum count, then rand will be always zero
                k = i Mod empColl.count
                Randomize
                'find the next min count
                'empsMin = empsEqual(1).count(0)
                'Set empsMinSecond = Nothing
                'Set empsMinSecond = findSecondMinOfClsPersons(empColl, empsMin)
                b_swap = True 'check if one employee doesn't have another shift day after
                If isHoliday Then
                    balanced = False
                Else
                    balanced = True
                End If
                If empsEqual.count = 0 Then
                        MsgBox "chyba"
                End If
                If empsEqual Is Nothing Then
                        MsgBox "chyba"
                End If
                Do While b_swap
                    dbgcntr2 = dbgcntr2 + 1
                    If dbgcntr2 = 3 * empColl.count Then 'to by mělo stačit na prostřídání všech zaměstnanců...
                        MsgBox "Selhání při pokusu o výpočet služby v pořadí " & d + 1 & " [while b_swap]"
                        'Exit Do
                    End If
                    Set emp = empsEqual(1)
                    empsMin = emp.count(params.shift_type + 1)
                    If isholiday Then empsMin = empsEqual(1).holidayWorks
                    If pickAnother And findnewmins Then: rand = Round((empsEqual.count - 1) * Rnd + 1)
                    b_swap = False
                    Dim rowoffset As Integer
                    If d < params.shiftInterval Then
                        rowoffset = d
                    Else
                        rowoffset = params.shiftInterval
                    End If
                    If params.perWeek Or Not findnewmins Then 'if calculate schedule is for whole work week, d needs to be shifted by 5 days (mon - fri)
                        'd = d + 4 - wd + 1 'important - final assigment of employee to work this shift
                        minIndex = empsEqual(rand).Id
                        arrIdx(k) = minIndex
                        If params.perWeek And findnewmins Then
                            findnewmins = False  'findnewmins na false, aby se generovalo jen v pondělí (nebo po pondělí, díky podmínce nahoře by mělo být true)
                        Else
                            If Not targetcells.Offset(d + 1, 0).Interior.ColorIndex = 38 And _
                            wd <> 5 Then
                                findnewmins = True 'jsou dva (a více) svátky po sobě?
                            End If
                            If wd <> 6 Then: balanced = False
                        End If
                    Else 'BIG OTHERWISE
                    If Not pickAnother Then: pickAnother = True  'Přesunuto z pod while
                    If Not display Then: display = True
                    'if shiftinterval is not fullfilled, try another employee
                    For cnt = 1 To rowoffset
                        If (targetcells.Offset(d - cnt, 0).Value = empsEqual(rand).Name) Then
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
                    If (Not b_swap And params.wkndRule > 0 And Not params.lDepend Is Nothing) Then
                        Dim lookforward As Integer, lookforward2 As Integer, wd2 as Integer
                        ' If wd = 4 And targetcells.Offset(d + 1, 0).Interior.ColorIndex = 38 Then
                        '     lookforward = 4
                        ' ElseIf wd = 5 Then
                        '     lookforward = 3
                        ' ElseIf wd = 6 Then
                        '     lookforward = 2
                        ' Else
                        '     lookforward = 1
                        ' End If
                        If isHoliday Then
                            lookforward = 0
                        Else
                            lookforward = 2 'standard pro sobotu
                        End If
                        'kontrola na pondělí, jestli je svátek, pak je třeba se koukat na úterý
                        wd2 = wd
                        While targetcells.Offset(d + lookforward, 0).Interior.ColorIndex = 38
                            If wd2 < 5 Then
                                lookforward = lookforward + 1
                                wd2 = wd2 + 1
                            Else
                                wd2 = (wd2 + 8-wd) Mod 7
                                lookforward = lookforward + 8-wd
                            End If
                        Wend
                        lookforward2 = lookforward + 7 'pro kontrolu o týden dále
                        'kontrola na pondělí o týden dále, jestli je svátek, pak je třeba se koukat na úterý. Podstatné jen pro |empsEqual.count = 2|!
                        wd2 = (wd + lookforward2) Mod 7 'na neděli by neměl nikdy koukat
                        While targetcells.Offset(d + lookforward2, 0).Interior.ColorIndex = 38
                            If wd2 < 5 Then
                                lookforward2 = lookforward2 + 1
                                wd2 = wd2 + 1
                            Else
                                wd2 = (wd2 + 8-wd) Mod 7
                                lookforward2 = lookforward2 + 8-wd
                            End If
                        Wend
                        ''''''''''''''''kontrola pauzy víkendů po sobě''''''''''''''''''''''
                        If wd <> 1 Then
                            Dim daysToWk as Integer
                            daysToWk = 0
                            If isHoliday and (lookforward > 2) Then
                                daysToWk = 6 - wd
                            End If
                        For ii = 1 To noWeeksBefore '2 'zde je možné regulovat "citlivost"
                            For jj = 0 To params.lDepend.Columns.count - 1
                                If (params.lDepend.Cells(1, 1).Offset(d - ii * 7 + (daysToWk), jj + 1).Value = empsEqual(rand).Name) Or _
                                (params.lDepend.Cells(1, 1).Offset(d + ii * 7 + (daysToWk), jj + 1).Value = empsEqual(rand).Name) Then 'pozor, to jj+1 je choulostivé podle zadání v parametrech.
                                    b_swap = True
                                    balanced = False
                                    unique = False
                                    notPassedWknd1 = True
                                    empsEqual.Remove rand
                                    Exit For
                                End If
                            Next
                            If b_swap Then: Exit For
                        Next
                        End If
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        If (Not b_swap And params.wkndRule = 1) Then
                            If (params.lDepend.Cells(1, 1).Offset(d + lookforward, 0).Value = empsEqual(rand).Name) Then
                                b_swap = True
                                empsEqual.Remove rand
                            ElseIf empsEqual.count = 2 Then 'check if the other employee with the same minimum count does not meet the condition for next week (thus d + 2 + 7) (lookforward2)
                                If (params.lDepend.Cells(1, 1).Offset(d + lookforward2, 0).Value = empsEqual(3 - rand).Name) Then '...then try this employe instead of current one if he meets all conditions for this week also.
                                    pickAnother = False
                                    rand = 3 - rand
                                    b_swap = True
                                    'balanced = False
                                End If
                            End If
                        ElseIf (Not b_swap And params.wkndRule = 2) Then
                            If (params.lDepend.Cells(1, 1).Offset(d + lookforward, 0).Value = empsEqual(rand).Name) Or _
                            (params.lDepend.Cells(1, 1).Offset(d + lookforward, 1).Value = empsEqual(rand).Name) Then
                                b_swap = True
                                empsEqual.Remove rand
                            ElseIf empsEqual.count = 2 Then 'check if the other employee with the same minimum count does not meet the condition for next week (thus d + 2 + 7)
                                If (params.lDepend.Cells(1, 1).Offset(d + lookforward2, 0).Value = empsEqual(3 - rand).Name) Or _
                                (params.lDepend.Cells(1, 1).Offset(d + lookforward2, 1).Value = empsEqual(3 - rand).Name) Then '...then try this employe instead of current one if he meets all conditions for this week also..
                                    pickAnother = False
                                    'noWeeksBefore = 0 '!toto je možnost - když v této situaci ten druhý nesplňuje podm. "po dvou týdnech" teď, ale za týden nemůže, můžeme podm. tímto zrušit a upřednostnit jeho volbu...
                                    rand = 3 - rand
                                    b_swap = True
                                    'balanced = False
                                End If
                            End If
                        End If
                        If Not b_swap Then
                            'If wd = 4  Then: findnewmins = False 'And targetcells.Offset(d + 1, 0).Interior.ColorIndex = 38 Then: findnewmins = False
                            If targetcells.Offset(d + 1, 0).Interior.ColorIndex = 38 Then: findnewmins = False 'jsou dva (a více) svátky po sobě?
                            If wd = 5 Then: findnewmins = False
                            If wd <> 6 Then: balanced = False
                        End If
                    End If
                    'check the noDayOfWeelRepeat rule - employee must not work the shift always in the same day of the week
                    'the "balanced" rule solves the case, where number of employees is a multiply of 5 - the schedule then must be random
                    If (Not b_swap And params.noDayOfWeekRepeat And (d >= weeksLookBack * 7)) Then
                        If (targetcells.Offset(d - (weeksLookBack * 7), 0).Value = empsEqual(rand).Name) Then
                            b_swap = True
                            balanced = False
                            empsEqual.Remove rand
                        End If
                    End If
                    ''''''''''''''''''''''''''''''''''''''''''
                    If Not b_swap Then: notPassedWknd1 = False
                    If Not b_swap Then: noWeeksBefore = 2
                    ''''''''''''''''''''''''''''''''''''''''''
                    'if employees with min are depleted (didn't match the requirements above), pick another set of employees find next lowest shift count
                    If b_swap And empsEqual.count = 0 Then
                        If noWeeksBefore > 0 And notPassedWknd1 Then
                            Set empsEqual = findMinOfClsPersons(empColl, params.shift_type + 1, isholiday)
                            noWeeksBefore = noWeeksBefore - 1
                        Else
                            Set empsEqual = findSecondMinOfClsPersons(empColl, empsMin, params.shift_type + 1, isholiday)
                        End If
                        balanced = False 'spoléháme se na shiftinterval < empColl.count. Díky této podmínce se ale může stát, že index je takový, že arrIdx
                        'se stane unikátním. Pak nechceme, aby se unikátnost kontrolovala a tedy byla stále false. Cyklus se pak bude stále snažit vybalancovat.
                        If Not empsEqual Is Nothing And empsEqual.count > 0 Then
                            Dim newmin
                            If isholiday Then
                                newmin = empsEqual(1).holidayWorks
                            Else
                                newmin = empsEqual(1).count(params.shift_type + 1)
                            End If
                            If newmin = empsMin Then 'Zde je snaha aby při nemožnosti splnění všech požadavků program neskončil v nekonečném cyklu.
                                'minIndex = -1 'make arridx position at i invalid and reassign in the next cycle
                                balanced = False
                                dbg = False
                                If dbg then
                                    MsgBox "Chyba: pro datum " & mydate & " se nepodařilo splnit všechny podmínky." & vbNewLine & "(Err1, noWeeksBefore = " & noWeeksBefore & ")"
                                Else
                                    MsgBox "Chyba: pro datum " & mydate & " se nepodařilo splnit všechny podmínky."
                                End If
                                targetcells.Offset(d, 0).Select
                                ' If noWeeksBefore > 0 Then
                                '     noWeeksBefore = noWeeksBefore - 1
                                '     'i = i - 1
                                '     'd = d - 1
                                ' End If
                                'Exit Do
                            End If
                        Else
                            MsgBox "Chyba: pro datum " & mydate & " se nepodařilo splnit všechny podmínky. (Err2)"
                            targetcells.Offset(d, 0).Select
                            findnewmins = True 'bude chyba, v po se nastaví stejně na false...
                            balanced = False
                            'i = i - 1
                            'd = d - 1
                            'If noWeeksBefore > 0 Then: noWeeksBefore = noWeeksBefore - 1 'TOHLE JE DIVNÝ
                            Exit Do
                        End If
                    ElseIf Not b_swap Then 'important - final assigment of employee to work this shift
                        minIndex = empsEqual(rand).Id
                        arrIdx(k) = minIndex
                    End If
                    'ENDIF OF BIG OTHERWISE
                    End If
                Loop
            If b_swap Then: display = False
            Else
                ' If params.perWeek Then 'if calculate schedule is for whole work week, d needs to be shifted by 5 days (mon - fri)
                '     d = d + 4 - wd + 1 'important - final assigment of employee to work this shift
                ' End If
                k = i Mod empColl.count
                'If Not targetcells.Offset(d - 1, 0).Interior.ColorIndex = 38 And
                If (Not isholiday And wd < 6) Or ((wd = 6 Or (isholiday And Not wd = 1)) And Not targetcells.Offset(d - 1, 0).Interior.ColorIndex = 38) Then
                    minIndex = arrIdx(k)' - holcnt) 'zde pozor, pokud je svátek, minIndex se neupdatuje, což je dobré řešení nyní, ale v případě změny - o svátcích nevypisovat každý den to může být problém a nutno vyřešit úplně jinak
                Else
                    'holcnt = holcnt + 1
                    i = i - 1 'minindex se neupdatnul, takže potřebujeme vrátit zpět i pořadí
                End If
            End If
            'if unique is true, we need to re-check if conditions are met
            If Not b_swap And display Then 'also b_swap if |exit do|!
            For Each emp In empColl
                If emp.Id = minIndex Then
                    b_swap = False 'warning! value was previously inversed, result is now the same, but the meaning wasn't correct before
                    If unique And Not isholiday Then
                    For cnt = 1 To rowoffset
                        'check again, if shiftInterval is fullfilled etc.
                        If (targetcells.Offset(d - cnt, 0).Value = emp.Name) Then
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
                    'Ošetření pravidel víkendu v případě "opakování kolečka zaměstnanců"
                    If (Not b_swap And params.wkndRule > 0 And Not params.lDepend Is Nothing) Then
                        Dim lookfwd As Integer
                        If wd = 4 And targetcells.Offset(d + 1, 0).Interior.ColorIndex = 38 Then
                            lookfwd = 4
                        ElseIf wd = 5 Then 'to tady teď nefunguje
                            lookfwd = 3
                        ElseIf wd = 6 Then
                            If targetcells.Offset(d + 2, 0).Interior.ColorIndex = 38 Then
                                lookfwd = 3
                            Else
                                lookfwd = 2
                            End If
                        Else
                            lookfwd = 1
                        End If
                        If (Not b_swap And params.wkndRule = 1 And Not params.lDepend Is Nothing) Then
                            If (params.lDepend.Cells(1, 1).Offset(d + lookfwd, 0).Value = emp.Name) Then
                                b_swap = True
                                balanced = False
                                unique = False
                            End If
                        End If
                        If (Not b_swap And params.wkndRule = 2 And Not params.lDepend Is Nothing) Then
                            If (params.lDepend.Cells(1, 1).Offset(d + lookfwd, 0).Value = emp.Name) Or _
                            (params.lDepend.Cells(1, 1).Offset(d + lookfwd, 1).Value = emp.Name) Then
                                b_swap = True
                                balanced = False
                                unique = False
                            End If
                        End If
                    End If
                    End If
                    If Not b_swap Then
                        'target here is a cell selected by us, recently "C10"
                        ' If params.perWeek Then
                        '     Dim enddays As Integer
                        '     enddays = 0
                        '     If d > daysinyear - 1 Then: enddays = d - daysinyear + 1
                        '     emp.count(0) = emp.count(0) + 5 - wd + 1 - enddays
                        '     emp.count(params.shift_type + 1) = emp.count(params.shift_type + 1) + 5 - wd + 1 - enddays
                        '     For Z = d - (4 - wd + 1) To d - enddays
                        '     targetcells.Offset(Z, 0).Value = emp.Name
                        '     targetcells.Offset(Z, 0).Font.Name = "Times New Roman"
                        '     Next
                        '     If i Mod empColl.count = 0 Then
                        '     targetcells.Offset(d - 4 - wd + 1, 0).Font.Bold = True
                        '     End If
                        ' Else
                        If isholiday And ((params.wkndRule > 0 And params.wkndRule <= 2) Or wd <> 1) Then
                            emp.holidayWorks = emp.holidayWorks + 1
                            emp.count(0) = emp.count(0) + 1
                        Else
                            emp.count(0) = emp.count(0) + 1
                            emp.count(params.shift_type + 1) = emp.count(params.shift_type + 1) + 1
                        End If
                        If Not isholiday Or (isholiday And ((params.wkndRule > 0 And params.wkndRule <= 2) Or wd <> 1)) Then
                            targetcells.Offset(d, 0).Value = emp.Name
                            targetcells.Offset(d, 0).Font.Name = "Times New Roman"
                            targetcells.Offset(d, 0).Font.Size = 12
                            ' If params.shift_type > 1 Then
                            '     targetcells.Offset(d, 0).Interior.Color = RGB(12, 192, 255)
                            ' End If
                            ' If i Mod empColl.count = 0 Then
                            '     targetcells.Offset(d, 0).Font.Bold = True
                            ' End If
                        End If
                        emplist.Cells(2, minIndex + 1).Value = emp.count(0)
                        i = i + 1
                    Else
                        'd = d - 1
                        increment = False
                        debugCounter = debugCounter + 1
                        If debugCounter = 366 Then
                            MsgBox "Selhání při pokusu o výpočet služby v pořadí " & d + 1 & " [while main]"
                            Exit Sub
                        End If
                    End If
                    Exit For
                End If
            Next
            End If
        End If
        End If
        'POST PROCESSING!!! (MOVE to separate procedure when possible)
        With targetcells.Offset(d, 0).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        If params.wkndRule = 4 Then
            With targetcells.Offset(d, 0).Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                 .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
        End If
        If wd > 5 Then
            'targetcells.Offset(d, 0).Interior.Color = RGB(12, 192, 255)
            If Not isholiday Then: targetcells.Offset(d, 0).Interior.ColorIndex = 37
            If wd = 6 Then
                With targetcells.Offset(d, 0).Borders(xlEdgeTop)
                 .LineStyle = xlContinuous
                    .Weight = xlThin 'xlMedium
                    .ColorIndex = xlAutomatic
                End With
            ElseIf wd = 7 Then
                With targetcells.Offset(d, 0).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                 .Weight = xlThin 'xlMedium
                    .ColorIndex = xlAutomatic
                End With
            End If
            ' With targetcells.Offset(d, 0).Borders(xlEdgeLeft)
            '     .LineStyle = xlContinuous
            '     .Weight = xlMedium
            '     .ColorIndex = xlAutomatic
            ' End With
            If params.wkndRule = 4 Then
                With targetcells.Offset(d, 0).Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlThin 'xlMedium
                    .ColorIndex = xlAutomatic
                End With
            End If
            If params.wkndRule > 0 And (params.wkndRule < 3) Then 'toto je zde jako rychlé dočasné řešení, které služby se mají podtrhnout a být tučně a které mají být jen tučně a které mají být "normálně"
                targetcells.Offset(d, 0).Font.Underline = True
                targetcells.Offset(d, 0).Font.Bold = True
            ElseIf params.wkndRule = 3 Then 'toto je zde jako rychlé dočasné řešení, které služby se mají podtrhnout a být tučně a které mají být jen tučně a které mají být "normálně"
                targetcells.Offset(d, 0).Font.Bold = True
            End If
        End If
        'END POST PROCESSING!!!
        If increment Then: d = d + 1
        mydate = CDate(params.target.Offset(d, 0).Value2)
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
        If holpos > 0 Then
            emplist.Cells(holpos, emp.Id + 1).ClearContents
            emplist.Cells(holpos, emp.Id + 1).Value = emp.holidayWorks
        End If
    Next
    'markHolidays
End Sub