Attribute VB_Name = "Module1"
Option Compare Text
Public dateString As String
Public dateStringFrom As String
Public dateStringTo As String
Public holpos As Integer
Public EmpCount As Integer

Type typeParams
    target As Range
    soffset As Integer
    dayfrom As Integer
    dayto As Integer
    shiftinterval As Integer
    lDepend As Range
    noDayOfWeekRepeat As Boolean
    noDayBefore As Boolean
    noDayAfter As Boolean
    wkndRule As Integer 'tady to pozmìnit
    perWeek As Boolean
    shift_type As Integer
End Type

Function checkGlobalVars() As Integer
    If dateString = "" Then
        MsgBox "Vyplòte/aktualizujte prosím buòku s datumem."
        checkGlobalVars = 1
    ElseIf dateStringFrom = "" Or dateStringTo = "" Then
        MsgBox "Vyplòte/aktualizujte prosím buòku s datumem ""od:""/""do:""."
        checkGlobalVars = 1
    Else
        checkGlobalVars = 0
    End If
End Function


Function LoadEmployees(ByVal target As Range) As Collection
    Dim empCol As New Collection
    Dim emp As clsPerson
    'Dim holpos As Integer
    holpos = 0
    
    Dim rg As Range
    Set rg = target.CurrentRegion
    For j = 1 To rg.Rows.count
        'If InStr(rg.Cells(j, 1).Value2, "svát") > 0 Then
        If rg.Cells(j, 1).Value Like "*svát*" Then
            holpos = j
            Exit For
        End If
    Next
    For i = 2 To rg.Columns.count
        If Not IsEmpty(rg.Cells(1, i)) Then
        Set emp = New clsPerson
        emp.Name = rg.Cells(1, i)
        emp.count(0) = rg.Cells(2, i)
        emp.count(1) = rg.Cells(3, i)
        emp.count(2) = rg.Cells(4, i)
        emp.count(3) = rg.Cells(5, i)
        emp.count(4) = rg.Cells(6, i)
        emp.Id = i - 1
        If holpos > 0 Then: emp.holidayWorks = rg.Cells(holpos, i)
        empCol.Add emp
        End If
    Next
    EmpCount = empCol.count
    Set LoadEmployees = empCol
End Function
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
    'Set datecell = Range("B10") 'možná lepší pøedávat parametrem
    'Set startDt = Range("I9") 'možná lepší pøedávat parametrem
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
    weeksLookBack = empColl.count \ 5 'pomocná proménná - pokud je poèet zamìstnancù násobek 5ti,
    'musíme nìjak zjistit, kolik týdnù zpìt se podívat, jestli tam zamìstnanec nemá službu.
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
    
    If params.shiftinterval >= empColl.count And params.wkndRule = 0 Then
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
    Dim wd As Integer, debugCounter As Integer, dbgcntr2 As Integer
    Dim pickAnother As Boolean: pickAnother = True
    Dim findnewmins As Boolean: findnewmins = True
    Dim iwantcontinuebuticant As Boolean: iwantcontinuebuticant = True 'já chci, ne program... není pøíkaz continue...
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
        If Not findnewmins And (((params.perWeek And wd = 1) Or Not params.perWeek)) And (params.wkndRule = 0) Then: findnewmins = True  'Zde by to mìlo fungovat správnì a zaruèit generování když je opravdu potøeba
        If Not findnewmins And (Not isholiday And params.wkndRule > 0 And wd = 5) Then: findnewmins = True 'kontrola - pøed víkendem chceme generovat nové, aby na sobotì nezùstalo generování vyp.
        'zde se postaráme, aby v létì o prázdninách a o vánocích nebyla služba 6-14:30, v pøípadì této služby od po-pá (perWeek)
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
        'If d Mod 7 < params.dayto And d Mod 7 >= params.dayfrom Then 'dayto = do jakého dne pojedeme, dayfrom = od jakého dne zaèneme ''STARÉ 'podmínka rozšíøena o kontrolu svátku podle barvy pozadí buòky...
        If (wd >= params.dayfrom And wd <= params.dayto And Not isholiday) Or _
        (wd <> 7 And params.wkndRule > 0 And isholiday) Then
            'pokud je svátek v pondìlí, nechceme generovat. Pracují lidé z víkendu.
            If params.wkndRule > 0 And isholiday And wd = 1 Then
                findnewmins = False
            End If
            If isholiday Then
                ' balanced = False
                ' unique = True 'HACK
                If Not balanced Then: unique = False  'tohle je tøeba ještì promyslet
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
                'If Not findnewmins And ((params.perWeek And wd = 1) Or Not params.perWeek) Then : findnewmins = True 'Asi to je potøeba... 'toto tady nefunguje správnì, pokud pondìlí zrovna "vynecháme z hlavního cyklu"
                If findnewmins And pickAnother Then 'Or empsEqual.count = 0 Then 'toto není dobrý fix...empsEqual by v tomto kroku nikdy nemìl být prázdný
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
                If isholiday Then
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
                    If dbgcntr2 = 3 * empColl.count Then 'to by mìlo staèit na prostøídání všech zamìstnancù...
                        MsgBox "Selhání pøi pokusu o výpoèet služby v poøadí " & d + 1 & " [while b_swap]"
                        'Exit Do
                    End If
                    Set emp = empsEqual(1)
                    empsMin = emp.count(params.shift_type + 1)
                    If isholiday Then empsMin = empsEqual(1).holidayWorks
                    If pickAnother And findnewmins Then: rand = Round((empsEqual.count - 1) * Rnd + 1)
                    b_swap = False
                    Dim rowoffset As Integer
                    If d < params.shiftinterval Then
                        rowoffset = d
                    Else
                        rowoffset = params.shiftinterval
                    End If
                    If params.perWeek Or Not findnewmins Then 'if calculate schedule is for whole work week, d needs to be shifted by 5 days (mon - fri)
                        'd = d + 4 - wd + 1 'important - final assigment of employee to work this shift
                        minIndex = empsEqual(rand).Id
                        arrIdx(k) = minIndex
                        If params.perWeek And findnewmins Then
                            findnewmins = False  'findnewmins na false, aby se generovalo jen v pondìlí (nebo po pondìlí, díky podmínce nahoøe by mìlo být true)
                        Else
                            If Not targetcells.Offset(d + 1, 0).Interior.ColorIndex = 38 And _
                            wd <> 5 Then
                                findnewmins = True 'jsou dva (a více) svátky po sobì?
                            End If
                            If wd <> 6 Then: balanced = False
                        End If
                    Else 'BIG OTHERWISE
                    If Not pickAnother Then: pickAnother = True  'Pøesunuto z pod while
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
                        Dim lookforward As Integer, lookforward2 As Integer, wd2 As Integer
                        ' If wd = 4 And targetcells.Offset(d + 1, 0).Interior.ColorIndex = 38 Then
                        '     lookforward = 4
                        ' ElseIf wd = 5 Then
                        '     lookforward = 3
                        ' ElseIf wd = 6 Then
                        '     lookforward = 2
                        ' Else
                        '     lookforward = 1
                        ' End If
                        If isholiday Then
                            lookforward = 0
                        Else
                            lookforward = 2 'standard pro sobotu
                        End If
                        'kontrola na pondìlí, jestli je svátek, pak je tøeba se koukat na úterý
                        wd2 = wd
                        While targetcells.Offset(d + lookforward, 0).Interior.ColorIndex = 38
                            If wd2 < 5 Then
                                lookforward = lookforward + 1
                                wd2 = wd2 + 1
                            Else
                                wd2 = (wd2 + 8 - wd) Mod 7
                                lookforward = lookforward + 8 - wd
                            End If
                        Wend
                        lookforward2 = lookforward + 7 'pro kontrolu o týden dále
                        'kontrola na pondìlí o týden dále, jestli je svátek, pak je tøeba se koukat na úterý. Podstatné jen pro |empsEqual.count = 2|!
                        wd2 = (wd + lookforward2) Mod 7 'na nedìli by nemìl nikdy koukat
                        While targetcells.Offset(d + lookforward2, 0).Interior.ColorIndex = 38
                            If wd2 < 5 Then
                                lookforward2 = lookforward2 + 1
                                wd2 = wd2 + 1
                            Else
                                wd2 = (wd2 + 8 - wd) Mod 7
                                lookforward2 = lookforward2 + 8 - wd
                            End If
                        Wend
                        ''''''''''''''''kontrola pauzy víkendù po sobì''''''''''''''''''''''
                        If wd <> 1 Then
                            Dim daysToWk As Integer
                            daysToWk = 0
                            If isholiday And (lookforward > 2) Then
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
                                    'noWeeksBefore = 0 '!toto je možnost - když v této situaci ten druhý nesplòuje podm. "po dvou týdnech" teï, ale za týden nemùže, mùžeme podm. tímto zrušit a upøednostnit jeho volbu...
                                    rand = 3 - rand
                                    b_swap = True
                                    'balanced = False
                                End If
                            End If
                        End If
                        If Not b_swap Then
                            'If wd = 4  Then: findnewmins = False 'And targetcells.Offset(d + 1, 0).Interior.ColorIndex = 38 Then: findnewmins = False
                            If targetcells.Offset(d + 1, 0).Interior.ColorIndex = 38 Then: findnewmins = False 'jsou dva (a více) svátky po sobì?
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
                        balanced = False 'spoléháme se na shiftinterval < empColl.count. Díky této podmínce se ale mùže stát, že index je takový, že arrIdx
                        'se stane unikátním. Pak nechceme, aby se unikátnost kontrolovala a tedy byla stále false. Cyklus se pak bude stále snažit vybalancovat.
                        If Not empsEqual Is Nothing And empsEqual.count > 0 Then
                            Dim newmin
                            If isholiday Then
                                newmin = empsEqual(1).holidayWorks
                            Else
                                newmin = empsEqual(1).count(params.shift_type + 1)
                            End If
                            If newmin = empsMin Then 'Zde je snaha aby pøi nemožnosti splnìní všech požadavkù program neskonèil v nekoneèném cyklu.
                                'minIndex = -1 'make arridx position at i invalid and reassign in the next cycle
                                balanced = False
                                dbg = False
                                If dbg Then
                                    MsgBox "Chyba: pro datum " & mydate & " se nepodaøilo splnit všechny podmínky." & vbNewLine & "(Err1, noWeeksBefore = " & noWeeksBefore & ")"
                                Else
                                    MsgBox "Chyba: pro datum " & mydate & " se nepodaøilo splnit všechny podmínky."
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
                            MsgBox "Chyba: pro datum " & mydate & " se nepodaøilo splnit všechny podmínky. (Err2)"
                            targetcells.Offset(d, 0).Select
                            findnewmins = True 'bude chyba, v po se nastaví stejnì na false...
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
                    minIndex = arrIdx(k) ' - holcnt) 'zde pozor, pokud je svátek, minIndex se neupdatuje, což je dobré øešení nyní, ale v pøípadì zmìny - o svátcích nevypisovat každý den to mùže být problém a nutno vyøešit úplnì jinak
                Else
                    'holcnt = holcnt + 1
                    i = i - 1 'minindex se neupdatnul, takže potøebujeme vrátit zpìt i poøadí
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
                    'Ošetøení pravidel víkendu v pøípadì "opakování koleèka zamìstnancù"
                    If (Not b_swap And params.wkndRule > 0 And Not params.lDepend Is Nothing) Then
                        Dim lookfwd As Integer
                        If wd = 4 And targetcells.Offset(d + 1, 0).Interior.ColorIndex = 38 Then
                            lookfwd = 4
                        ElseIf wd = 5 Then 'to tady teï nefunguje
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
                            MsgBox "Selhání pøi pokusu o výpoèet služby v poøadí " & d + 1 & " [while main]"
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
            If params.wkndRule > 0 And (params.wkndRule < 3) Then 'toto je zde jako rychlé doèasné øešení, které služby se mají podtrhnout a být tuènì a které mají být jen tuènì a které mají být "normálnì"
                targetcells.Offset(d, 0).Font.Underline = True
                targetcells.Offset(d, 0).Font.Bold = True
            ElseIf params.wkndRule = 3 Then 'toto je zde jako rychlé doèasné øešení, které služby se mají podtrhnout a být tuènì a které mají být jen tuènì a které mají být "normálnì"
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
Function findMinOfClsPersons(ByRef iCol As Collection, shifttype As Integer, isholiday As Boolean) As Collection
    Dim emp As clsPerson, minemps As New Collection, empsMin As Long
    If isholiday Then
        empsMin = iCol(1).holidayWorks
        For Each emp In iCol
            If emp.holidayWorks < empsMin Then
                empsMin = emp.holidayWorks
                Set minemps = Nothing
                minemps.Add emp
            ElseIf emp.holidayWorks = empsMin Then
            'if any other employee has the same count, we choose randomly, who gets the next shift
                minemps.Add emp
            End If
        Next
    Else
        empsMin = iCol(1).count(shifttype)
        For Each emp In iCol
            If emp.count(shifttype) < empsMin Then
                empsMin = emp.count(shifttype)
                Set minemps = Nothing
                minemps.Add emp
            ElseIf emp.count(shifttype) = empsMin Then
            'if any other employee has the same count, we choose randomly, who gets the next shift
                minemps.Add emp
            End If
        Next
    End If
    Set findMinOfClsPersons = minemps
End Function

Function findSecondMinOfClsPersons(ByRef iCol As Collection, firstmin As Long, shifttype As Integer, isholiday As Boolean) As Collection
    Dim emp As clsPerson, minemps As New Collection, empsMin As Long
    empsMin = firstmin
    If isholiday Then
       For Each emp In iCol
            If empsMin > firstmin And firstmin < emp.holidayWorks And emp.holidayWorks < empsMin Then
                empsMin = emp.holidayWorks
                Set minemps = Nothing
                minemps.Add emp
            ElseIf empsMin = firstmin And emp.holidayWorks > empsMin Then
                empsMin = emp.holidayWorks
                minemps.Add emp
            ElseIf emp.holidayWorks = empsMin And empsMin >= firstmin Then
                minemps.Add emp
            End If
        Next
    Else
        For Each emp In iCol
            If empsMin > firstmin And firstmin < emp.count(shifttype) And emp.count(shifttype) < empsMin Then
                empsMin = emp.count(shifttype)
                Set minemps = Nothing
                minemps.Add emp
            ElseIf empsMin = firstmin And emp.count(shifttype) > empsMin Then
                empsMin = emp.count(shifttype)
                minemps.Add emp
            ElseIf emp.count(shifttype) = empsMin And empsMin >= firstmin Then
                minemps.Add emp
            End If
        Next
    End If
    'If minemps.count = 0 Then
    '    empsMin = empsMin
    'End If
    Set findSecondMinOfClsPersons = minemps
End Function

Function checkForErrors(ByVal target As Range, shiftinterval As Integer, errfound As Boolean) As Boolean
    Dim rg As Range
    Set rg = target.CurrentRegion
    For i = 1 To rg.Rows.count
        For j = 1 To 2 'rg.Columns.Count
            If target.Cells(i, j).Value = target.Cells(i, j + 1).Value And _
            Not IsEmpty(target.Cells(i, j).Value) Then
                target.Cells(i, j).Interior.Color = RGB(255, 32, 32)
                target.Cells(i, j + 1).Interior.Color = RGB(255, 0, 0)
                errfound = True
                target.Cells(i, j).Select
                'MsgBox "Chyba: " & target.Cells(i, j).Value & " má pøekrývající se služby." _
                '& " Øádek: " & i
            ElseIf target.Cells(i, j).Value = target.Cells(i - shiftinterval, j).Value And _
            Not IsEmpty(target.Cells(i, j).Value) Then
                target.Cells(i, j).Interior.Color = RGB(255, 0, 0)
                target.Cells(i - shiftinterval, j).Interior.Color = RGB(192, 127, 0)
                errfound = True
                target.Cells(i, j).Select
                'MsgBox "Chyba: " & target.Cells(i, j).Value & " nesplòuje podmínku alespoò " _
                '& shiftInterval & " mezi službami." & " Øádek: " & i & " ve sloupci " & target.Cells(0, j).Value
            ElseIf Not errfound Then
                target.Cells(i, j).Interior.Color = xlNone
                target.Cells(i, j + 1).Interior.Color = xlNone
                target.Cells(i - shiftinterval, j).Interior.Color = xlNone
            End If
        Next
    Next
    checkForErrors = errfound
End Function

Function checkForErrors2(ByVal target As Range, errfound As Boolean) As Boolean
    Dim rg As Range
    Set rg = target.CurrentRegion
    For d = 0 To rg.Rows.count
    For cnt = 1 To target.Columns.count
        If Not IsEmpty(target.Cells(1, 1).Offset(d, 0).Value) And (target.Cells(1, 1).Offset(d - 1, -cnt).Value = target.Cells(1, 1).Offset(d, 0).Value) Then
            target.Cells(1, 1).Offset(d, 0).Interior.Color = RGB(200, 128, 0)
            target.Cells(1, 1).Offset(d - 1, -cnt).Interior.Color = RGB(200, 128, 0)
            errfound = True
            target.Cells(1, 1).Offset(d, 0).Select
            'MsgBox "Chyba: " & target.Cells(1, 1).Offset(d, 0).Value & " nesplòuje podmínku " & _
            '"služeb po sobì na øádku " & d + 1
            'Exit For
        End If
        If Not IsEmpty(target.Cells(1, 1).Offset(d, 0).Value) And (target.Cells(1, 1).Offset(d + 1, -cnt).Value = target.Cells(1, 1).Offset(d, 0).Value) Then
            target.Cells(1, 1).Offset(d, 0).Interior.Color = RGB(200, 128, 0)
            target.Cells(1, 1).Offset(d + 1, -cnt).Interior.Color = RGB(200, 128, 0)
            errfound = True
            target.Cells(1, 1).Offset(d, 0).Select
            'MsgBox "Chyba: " & target.Cells(1, 1).Offset(d, 0).Value & " nesplòuje podmínku " & _
            '"služeb po sobì na øádku " & d + 1
            'Exit For
            End If
        If Not errfound Then
            target.Cells(1, 1).Offset(d, 0).Interior.Color = xlNone
            target.Cells(1, 1).Offset(d - 1, -cnt).Interior.Color = xlNone
            target.Cells(1, 1).Offset(d + 1, -cnt).Interior.Color = xlNone
        End If
    Next
    Next
    checkForErrors2 = errfound
End Function
Function checkForErrors3(ByVal target As Range, errfound As Boolean) As Boolean
    Dim rg As Range
    Set rg = target.CurrentRegion
    For d = 0 To rg.Rows.count
        If Not IsEmpty(target.Cells(1, 1).Offset(d, 0).Value) And (target.Cells(1, 1).Offset(d + 2, -2).Value = target.Cells(1, 1).Offset(d, 0).Value) Then
            target.Cells(1, 1).Offset(d, 0).Interior.Color = RGB(200, 128, 0)
            target.Cells(1, 1).Offset(d + 2, -2).Interior.Color = RGB(200, 128, 0)
            errfound = True
            target.Cells(1, 1).Offset(d, 0).Select
            'MsgBox "Chyba: " & target.Cells(1, 1).Offset(d, 0).Value & " nesplòuje podmínku " & _
            '"služeb po sobì na øádku " & d + 1
            'Exit For
        End If
        If Not errfound Then
            target.Cells(1, 1).Offset(d, 0).Interior.Color = xlNone
            target.Cells(1, 1).Offset(d + 2, -2).Interior.Color = xlNone
        End If
    Next
    checkForErrors3 = errfound
End Function

Sub ReCalc()
    Dim keycells As Range
    Dim rg As Range
    Dim rng As Range, cnt As Integer
    
    For l = 1 To 8
        'If InStr(rg.Cells(j, 1).Value2, "svát") > 0 Then
        If Range("A1").Cells(l, 1).Value2 Like "*svát*" Then
            holpos = l
            Exit For
        End If
    Next
    
    Dim ret As Integer
    ret = checkGlobalVars()
    If ret > 0 Then
        Exit Sub
    End If
    LoadEmployees Range("A1")
    Set rng = Range("B1:Z1")
    If holpos > 0 Then
        Set rg = Range("B2", Range("A1").Offset(holpos - 1, EmpCount)) 'pozor toto je závislé "natvrdo"
    Else
        Set rg = Range("B2", Range("A1").Offset(5, EmpCount))
    End If
    rg.ClearContents
    Set keycells = Range(dateString)
        Dim j As Integer: j = 0
        Do While Not IsEmpty(keycells.Offset(j, 0))
            For k = 1 To 5
                For Each cell In rng
                    If IsEmpty(cell) Then: Exit For  'pozor, tabulka zamìstnancù musí být souvislá...
                    If keycells.Offset(j, k).Value2 = cell.Value2 Then
                        'celkový poèet (je na prvním øádku)
                        cnt = cell.Offset(1, 0).Value2
                        cnt = cnt + 1
                        cell.Offset(1, 0).Value2 = cnt
                        'konkrétní poèet - nutno zvýšit k o 1
                        If keycells.Offset(j, k).Interior.ColorIndex = 38 Then
                            If holpos > 0 Then
                            cnt = cell.Offset(holpos - 1, 0).Value2
                            cnt = cnt + 1
                            cell.Offset(holpos - 1, 0).Value2 = cnt
                            End If
                        Else
                            cnt = cell.Offset(k + 1, 0).Value2
                            cnt = cnt + 1
                            cell.Offset(k + 1, 0).Value2 = cnt
                        End If
                        Exit For 'Pøedpokládáme, že v horní tabulce zamìstnance je každý zamìstnanec pouze jednou
                    End If
                    'If IsEmpty(cell) Then : Exit For 'pozor, tabulka zamìstnancù musí být souvislá...
                Next
            Next
            j = j + 1
        Loop
End Sub

Sub PrepareTable(target As Range)
    With target.Offset(1, 0)
        .Value = "Poèet služeb"
        .Font.ColorIndex = 1
        .Font.Name = "Times New Roman"
        .Font.Size = 12
        .Font.Bold = True
    End With
    'target.Offset(2, 0).Value = "Poèet 6-14:30h"
    With target.Offset(2, 0)
        .Value = "Poèet 6-18h"
        .Font.ColorIndex = 3
        .Font.Name = "Times New Roman"
        .Font.Size = 12
        .Font.Bold = True
    End With
    With target.Offset(3, 0)
        .Value = "Poèet So + Ne"
        .Font.ColorIndex = 46
        .Font.Name = "Times New Roman"
        .Font.Size = 12
        .Font.Bold = True
    End With
    With target.Offset(4, 0)
        .Value = "Poèet So"
        .Font.ColorIndex = 4
        .Font.Name = "Times New Roman"
        .Font.Size = 12
        .Font.Bold = True
    End With
    With target.Offset(5, 0)
        .Value = "Poèet So pøísl."
        .Font.ColorIndex = 33
        .Font.Name = "Times New Roman"
        .Font.Size = 12
        .Font.Bold = True
    End With
    With target.Offset(6, 0)
        .Value = "Služby o svátcích"
        .Font.ColorIndex = 38
        .Font.Name = "Times New Roman"
        .Font.Size = 12
        .Font.Bold = True
    End With
End Sub

Sub CalcDates(keycells As Range)
    Application.EnableEvents = False
    dateString = keycells.Address
    'MsgBox "change"
    'Dim keycells As Range

    ' The variable KeyCells contains the cells that will
    ' cause an alert when they are changed.
    ' Set keycells = Range("B10")
    Dim mydate As Date, endDate As Date
    mydate = CDate(keycells.Value2)
    
    Dim enddatestr As String
    Dim dtyear As Integer, daysinyear As Integer, wd As Integer
    dtyear = Year(mydate)
    enddatestr = dtyear & "/12/31"
    endDate = CDate(enddatestr)
    daysinyear = DateDiff("d", mydate, endDate) + 1
        MsgBox "Výpoèet zmìnìn na rok " & Year(mydate) & ". V tomto roce je " & daysinyear & " dnù."
        For i = 1 To 366
            keycells.Offset(i, 0).clear
        Next
        With keycells.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With keycells.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        For i = 0 To daysinyear - 1
            keycells.Offset(i, 0).Value = DateSerial(dtyear, Month(mydate), Day(mydate) + i)
            keycells.Offset(i, 0).Font.Name = "Times New Roman"
            keycells.Offset(i, 0).Font.Size = 12
            keycells.Offset(i, 0).NumberFormat = "dd.mm."
            keycells.Offset(i, 0).HorizontalAlignment = xlCenter
            
            wd = Weekday(DateSerial(dtyear, Month(mydate), Day(mydate) + i), vbMonday)

            With keycells.Offset(i, 0).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            With keycells.Offset(i, 0).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            If wd > 5 Then
                keycells.Offset(i, 0).Interior.ColorIndex = 37
            If wd = 6 Then
                With keycells.Offset(i, 0).Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlThin 'xlMedium
                    .ColorIndex = xlAutomatic
                End With
            ElseIf wd = 7 Then
                With keycells.Offset(i, 0).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin 'xlMedium
                    .ColorIndex = xlAutomatic
                End With
            End If
            With keycells.Offset(i, 0).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlThin 'xlMedium
                .ColorIndex = xlAutomatic
            End With
            ' With KeyCells.Offset(i, 0).Borders(xlEdgeRight)
            '     .LineStyle = xlContinuous
            '     .Weight = xlMedium
            '     .ColorIndex = xlAutomatic
            ' End With
            End If
        Next
    markHolidays
    Application.EnableEvents = True
End Sub
