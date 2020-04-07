Attribute VB_Name = "Module4"
Sub tlaèítko3_Kliknutí()
    Dim params618 As typeParams, params61430 As typeParams, paramsW As typeParams, paramsSat As typeParams, _
    paramsSat2 As typeParams
    
    With params61430
        Set .Target = Range("C10")
            .daystart = 0
            .numofdays = 5
            .shiftInterval = 0
        Set .lDepend = Nothing
            .noDayOfWeekRepeat = False
            .noDayBefore = False
            .noDayAfter = False
            .wkndRule = False
            .perWeek = True
    End With
    
    With params618
        Set .Target = Range("D10")
            .daystart = 0
            .numofdays = 5
            .shiftInterval = 7
        Set .lDepend = Range("C10")
            .noDayOfWeekRepeat = True
            .noDayBefore = False
            .noDayAfter = False
            .wkndRule = False
            .perWeek = False
    End With
    
    paramsW = params618
    With paramsW
        Set .Target = Range("E10")
            .daystart = 5
            .numofdays = 6
            .shiftInterval = 0
        Set .lDepend = Range("c10:d10")
            .noDayOfWeekRepeat = False
            .wkndRule = True
    End With
    
    paramsSat = paramsW
    With paramsSat
        Set .Target = Range("F10")
        Set .lDepend = Range("E10:C10")
            .wkndRule = False
    End With
    
    paramsSat2 = paramsSat
    With paramsSat2
        Set .Target = Range("G10")
        Set .lDepend = Range("F10:C10")
    End With
    
    calcSchedulePerWeek params61430 'od 6 do pul 3
    calcSchedule params618 '12ctky
    calcSchedulePerWeek paramsW 'sobota+nedele
    calcSchedule paramsSat 'jen sobota
    calcSchedule paramsSat2 'jen sobota prisluzba
End Sub
