Function LoadEmployees(ByVal target As Range) As Collection
    Dim empCol As New Collection
    Dim emp As clsPerson
    
    Dim rg As Range
    Set rg = target.CurrentRegion
    For i = 2 To rg.Columns.Count
        Set emp = New clsPerson
        emp.Name = rg.Cells(1, i)
        emp.Count = rg.Cells(2, i)
        emp.Id = i - 1
        
        empCol.Add emp
    Next
    Set LoadEmployees = empCol
End Function
Sub calcSchedule(ByVal target As Range)
    'Variables
    Dim emps(1 To 3) As Long 'emps = employees
    Dim empsMin As Long, minIndex As Long, rndval As Long
    Dim arrIdx() As Long
    Dim var As Variant
    Dim empColl As Collection, empsEqual As New Collection
    
    Dim emplist As Range
    'here we must specify the cell where table with employees starts, in our test case "O1"
    Set emplist = Range("A1")
    Set empColl = LoadEmployees(emplist) 'employees Collection
    ReDim arrIdx(empColl.Count - 1)
    'Code
    'Randomize
        'For i = 1 To empColl.Count
         '   rndval = Int(empColl.Count * Rnd + 1)
          '  var = Application.Match(rndval, arrIdx, 0)
           ' While Not IsError(var)
            '    rndval = Int(empColl.Count * Rnd + 1)
             '   var = Application.Match(rndval, arrIdx, 0)
            'Wend
           ' arrIdx(i) = rndval
        'Next
        
        For j = 0 To 364 'empColl.Count
            target.Offset(j, 0).ClearContents
        Next
        For i = 0 To 364 'empColl.Count
            If i < empColl.Count Then
                Dim emp As clsPerson
                Set empsEqual = findMinOfClsPersons(empColl)
                'if two or more employees have the same minimal shift count we choose randomly
                'who gets the next shift. If only one has the minimum count, then rand will be always zero
                Dim rand As Long
                Randomize
                rand = Round((empsEqual.Count - 1) * Rnd + 1)
                minIndex = empsEqual(rand).Id
                arrIdx(i) = minIndex
            Else
                Dim k As Integer
                k = i Mod empColl.Count
                minIndex = arrIdx(k)
            End If
                
            For Each emp In empColl
                If emp.Id = minIndex Then
                    emp.Count = emp.Count + 1
                    'target here is a cell selected by us, recently "C3"
                    target.Offset(i, 0).Value = emp.Name
                    If i Mod empColl.Count = 0 Then
                        target.Offset(i, 0).Font.Bold = True
                    End If
                    emplist.Cells(2, minIndex + 1).Value = emp.Count
                    Exit For
                End If
            Next
        Next
        For Each emp In empColl
            Dim idx As Variant
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
            emplist.Cells(3, emp.Id + 1).Clear
            emplist.Cells(3, emp.Id + 1).Value = idx
        Next
End Sub

Function findMinOfClsPersons(ByRef iCol As Collection) As Collection
    Dim emp As clsPerson, minemps As New Collection, empsMin As Long
    empsMin = iCol(1).Count
    For Each emp In iCol
        If emp.Count < empsMin Then
            empsMin = emp.Count
            Set minemps = Nothing
            minemps.Add emp
        ElseIf emp.Count = empsMin Then
        'if any other employee has the same count, we choose randomly, who gets the next shift
            minemps.Add emp
        End If
    Next
    Set findMinOfClsPersons = minemps
End Function
